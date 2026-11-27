# neoffice's PDF exports (see pdf.lua): one LibreOffice, started once and kept running, converts each
# document it is asked to, so an export takes a fraction of a second instead of LibreOffice's start.
#
#   python3 pdf_server.py PROFILE_DIR WORK_DIR
#
# Requests on stdin, one JSON object a line: {"id": n, "src": document, "dst": pdf}. Answers on
# stdout, one a line: {"id": n, "ok": true, "ms": time} or {"id": n, "ok": false, "error": why}.
# The document is converted from a copy in WORK_DIR (so LibreOffice's lock file isn't left next to
# it); the PDF is written next to `dst` under a hidden name, then put in its place, so a viewer never
# reads half of one. When stdin closes (Neovim quit), LibreOffice is stopped.
import json
import os
import secrets
import shutil
import signal
import subprocess
import sys
import threading
import time

try:
    import uno
    from com.sun.star.beans import PropertyValue
except ImportError as e:  # (no LibreOffice's Python here: neoffice converts another way)
    print(json.dumps({"id": 0, "ok": False, "error": "no-uno: " + str(e)}), flush=True)
    sys.exit(3)

PROFILE, WORK = sys.argv[1], sys.argv[2]
os.makedirs(WORK, exist_ok=True)
office = None  # the LibreOffice process
desktop = None


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


PIDFILE = os.path.join(WORK, "office.pid")


def stop_left_over():
    # (one a Neovim that crashed left running: its pid written down, and still neoffice's, by its profile)
    try:
        pid = int(open(PIDFILE).read().strip())
        with open("/proc/%d/cmdline" % pid, "rb") as f:
            if PROFILE.encode() in f.read():
                os.killpg(pid, signal.SIGKILL)
    except Exception:
        pass


def start():
    global office, desktop
    stop()
    stop_left_over()
    # (reached through a pipe of a name no one can guess, not a TCP port: a port anyone on the machine
    # could connect to, and through it run whatever they like as you)
    pipe = "neoffice_pdf_" + secrets.token_hex(16)
    office = subprocess.Popen(
        ["soffice", "-env:UserInstallation=" + uno.systemPathToFileUrl(PROFILE), "--headless", "--invisible",
         "--norestore", "--nologo", "--nodefault", "--nolockcheck",
         "--accept=pipe,name=%s;urp;" % pipe],
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        start_new_session=True,  # (a group of its own: all of it stopped at once, see stop)
        preexec_fn=lambda: os.umask(0o077))  # (its pipe, and what it writes, yours alone)
    try:
        with open(PIDFILE, "w") as f:
            f.write(str(office.pid))
    except Exception:
        pass
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    url = "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipe
    for _ in range(600):  # (up to 30 s: its first start makes its profile)
        try:
            ctx = resolver.resolve(url)
            desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            return
        except Exception:
            if office.poll() is not None:
                break
            time.sleep(0.05)
    raise RuntimeError("LibreOffice didn't start")


def stop():
    global office, desktop
    if desktop is not None:
        try:
            desktop.terminate()
        except Exception:
            pass
    desktop = None
    if office is not None:
        try:
            office.wait(5)
        except Exception:
            pass
        try:  # (and whatever of it is left: soffice starts soffice.bin)
            os.killpg(office.pid, signal.SIGKILL)
        except Exception:
            pass
        try:
            os.remove(PIDFILE)
        except Exception:
            pass
    office = None


def quit_now(*_):
    stop()
    os._exit(0)


# LibreOffice never outlives Neovim: stopped when this is told to stop, or when Neovim is gone
signal.signal(signal.SIGTERM, quit_now)
signal.signal(signal.SIGHUP, quit_now)
PARENT = os.getppid()


def watch_parent():
    while True:
        time.sleep(1)
        if os.getppid() != PARENT:
            quit_now()


threading.Thread(target=watch_parent, daemon=True).start()


def convert(src, dst):
    copy = os.path.join(WORK, "document" + os.path.splitext(src)[1].lower())
    shutil.copyfile(src, copy)
    part = os.path.join(os.path.dirname(dst) or ".", "." + os.path.basename(dst) + ".neoffice-part")
    # (no macros run, no links updated: a document from someone else does nothing but be drawn)
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(copy), "_blank", 0,
                                       (prop("Hidden", True), prop("ReadOnly", True),
                                        prop("MacroExecutionMode", 0), prop("UpdateDocMode", 0)))
    if doc is None:
        raise RuntimeError("LibreOffice couldn't open the document")
    try:
        doc.storeToURL(uno.systemPathToFileUrl(part), (prop("FilterName", "writer_pdf_Export"),))
        os.replace(part, dst)
    finally:
        try:
            doc.close(True)
        except Exception:
            pass
        for path in (part, copy):
            if os.path.exists(path):
                os.remove(path)


for line in sys.stdin:
    try:
        req = json.loads(line)
    except ValueError:
        continue
    t0 = time.time()
    try:
        if desktop is None or office is None or office.poll() is not None:
            start()
        try:
            convert(req["src"], req["dst"])
        except Exception:
            if office is not None and office.poll() is None:
                raise
            start()  # (LibreOffice went away: started again, once)
            convert(req["src"], req["dst"])
        answer = {"id": req["id"], "ok": True, "ms": int((time.time() - t0) * 1000)}
    except Exception as e:
        answer = {"id": req.get("id"), "ok": False, "error": str(e)}
    print(json.dumps(answer), flush=True)
stop()
