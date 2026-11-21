#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows quick installer for the LibreOffice MCP co-editing setup.

  1. Locates a LibreOffice install (well-known paths, PATH, registry, or a
     directory you pass on the command line).
  2. Creates a private working folder under the current user's home:
     C:\\Users\\<you>\\.libreoffice-mcp\\
  3. Writes a .env with the detected values (soffice path, UNO port, profile).
  4. Copies lo_mcp_server.py next to it (a fresh copy, so the folder is
     self-contained and survives later edits to this repository).
  5. Generates start_shared.py, a generic launcher that reads .env, starts
     LibreOffice with a UNO socket, and detaches it from the caller.
  6. Registers spreadsheet and document extensions (.xlsx .xls .ods .csv .tsv
     .xlsm .docx .doc .odt .rtf) per-user (HKCU), without admin rights. The
     open command is soffice.exe itself with the UNO socket arguments inlined,
     so a double-click both opens the file and enables MCP co-editing. Because
     soffice.exe is a GUI-subsystem binary there is no console window, and the
     ProgID carries a DefaultIcon pointing at the coloured application icon
     embedded in soffice.bin instead of the grey generic document icon.
  7. Writes mcp-config.json filled in with this machine's actual paths, ready
     to paste into any MCP client (Claude Desktop, Cursor, etc).

Usage:
    python scripts/install_windows.py                 # auto-detect LibreOffice
    python scripts/install_windows.py --soffice "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
    python scripts/install_windows.py --port 2003     # different UNO port
    python scripts/install_windows.py --no-assoc      # skip file associations
    python scripts/install_windows.py --no-copy       # skip server copy

Re-run any time; it is idempotent and just refreshes the generated files.
"""
import argparse
import os
import shutil
import subprocess
import sys
import winreg

# --------------------------------------------------------------------------
# where things live
# --------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)
DEFAULT_HOME = os.path.expanduser("~")
DATA_DIR = os.path.join(DEFAULT_HOME, ".libreoffice-mcp")
ENV_FILE = os.path.join(DATA_DIR, ".env")
SERVER_SRC = os.path.join(REPO_ROOT, "lo_mcp_server.py")

def find_python(soffice):
    """Pick the interpreter to run lo_mcp_server.py with.

    Prefer the Python bundled with LibreOffice (it can import uno), falling
    back to a system Python when that is not available.
    """
    candidates = []
    if soffice:
        prog = os.path.dirname(soffice)
        candidates.append(os.path.join(prog, "python.exe"))
        candidates.append(os.path.join(prog, "pythonw.exe"))
    for c in candidates:
        if os.path.isfile(c):
            return c
    for name in ("python.exe", "python"):
        found = shutil.which(name)
        if found:
            return os.path.abspath(found)
    return sys.executable

# extensions -> ProgID
CALC_EXTS = (".xlsx", ".xls", ".xlsm", ".ods", ".csv", ".tsv")
WRITER_EXTS = (".docx", ".doc", ".odt", ".rtf")
CALC_PROGID = "LibreOffice.MCP.Calc"
WRITER_PROGID = "LibreOffice.MCP.Writer"


def log(msg):
    print("[install] %s" % msg, flush=True)


# --------------------------------------------------------------------------
# LibreOffice discovery
# --------------------------------------------------------------------------

def _registry_soffice():
    """Ask the registry where LibreOffice was installed."""
    candidates = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\LibreOffice\UNO\InstallPath"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\LibreOffice\UNO\InstallPath"),
    ]
    for hive, key in candidates:
        try:
            k = winreg.OpenKey(hive, key)
            try:
                val, _ = winreg.QueryValueEx(k, None)
            finally:
                winreg.CloseKey(k)
            if val:
                p = os.path.join(val, "program", "soffice.exe")
                if os.path.isfile(p):
                    return p
        except OSError:
            continue
    return None


def _well_known_soffice():
    roots = [
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
        os.environ.get("LOCALAPPDATA", ""),
    ]
    for root in roots:
        if not root:
            continue
        for d in ("LibreOffice", "LibreOffice 6", "LibreOffice 7",
                  "LibreOffice 24", "LibreOffice 25", "LibreOffice 26"):
            p = os.path.join(root, d, "program", "soffice.exe")
            if os.path.isfile(p):
                return p
    return None


def _path_soffice():
    for name in ("soffice.exe", "soffice"):
        found = shutil.which(name)
        if found:
            return os.path.abspath(found)
    return None


def find_soffice(explicit=None):
    if explicit:
        p = os.path.abspath(explicit)
        if not os.path.isfile(p):
            sys.exit("no such file: %s" % p)
        return p
    for finder in (_registry_soffice, _well_known_soffice, _path_soffice):
        p = finder()
        if p:
            return p
    return None


# --------------------------------------------------------------------------
# generated files
# --------------------------------------------------------------------------

def write_env(soffice, port, profile):
    with open(ENV_FILE, "w", encoding="utf-8", newline="\n") as f:
        f.write("# LibreOffice MCP shared-instance launcher configuration\n")
        f.write("LO_MCP_SOFFICE=%s\n" % soffice)
        f.write("LO_MCP_PORT=%s\n" % port)
        f.write("LO_MCP_PROFILE=%s\n" % profile)
    log("wrote %s" % ENV_FILE)


def write_launcher(port, profile, pythonw):
    """start_shared.py: reads .env, starts soffice detached with UNO socket."""
    script = r'''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Start LibreOffice with a UNO socket so the MCP server can attach to it.
Reads LO_MCP_SOFFICE / LO_MCP_PORT / LO_MCP_PROFILE from .env (this folder).

Usage:
    python start_shared.py [file]     # open a blank Calc, or the given file
"""
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ENV = os.path.join(HERE, ".env")


def _fatal(msg):
    """Show the error on screen (never fail silently) and exit."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "LibreOffice MCP",
                                         0x10)  # MB_ICONERROR
    except Exception:
        print(msg, file=sys.stderr, flush=True)
    sys.exit(1)


def load_env():
    env = {}
    if os.path.isfile(ENV):
        for line in open(ENV, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, _, v = line.partition("=")
                env[k.strip()] = v.strip()
    return env


def port_listening(port):
    r = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "if (Get-NetTCPConnection -LocalPort %s -State Listen "
         "-ErrorAction SilentlyContinue) { 'yes' } else { 'no' }" % port],
        capture_output=True, text=True)
    return r.stdout.strip() == "yes"


def main():
    env = load_env()
    soffice = env.get("LO_MCP_SOFFICE", "")
    port = env.get("LO_MCP_PORT", "2002")
    profile = env.get("LO_MCP_PROFILE", "")

    if not soffice or not os.path.isfile(soffice):
        _fatal("soffice not configured; run install_windows.py first")
    os.makedirs(profile, exist_ok=True)
    profile_url = "file:///" + profile.replace("\\", "/").lstrip("/")

    cmd = [
        soffice,
        "-env:UserInstallation=%s" % profile_url,
        "--nodefault", "--nologo", "--nolockcheck", "--norestore",
        "--nofirststartwizard",
        "--accept=socket,host=127.0.0.1,port=%s;urp;"
        "StarOffice.ComponentContext" % port,
    ]
    if len(sys.argv) > 1:
        target = os.path.abspath(sys.argv[1])
        if not os.path.isfile(target):
            _fatal("no such file: %s" % target)
        cmd.append(target)
    else:
        cmd.append("--calc")

    # Detach so LibreOffice survives this script exiting.
    quoted = " ".join('"%s"' % c if (" " in c or ";" in c) else c for c in cmd)
    ps = ("$r = Invoke-CimMethod -ClassName Win32_Process "
          "-MethodName Create -Arguments @{ CommandLine = '%s' }; "
          "Write-Output $r.ReturnValue" % quoted.replace("'", "''"))
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, text=True)
    if r.stdout.strip() != "0":
        _fatal("failed to start LibreOffice: %s" % r.stderr[:300])

    for _ in range(40):
        time.sleep(1)
        if port_listening(port):
            print("LibreOffice up, UNO port %s ready" % port)
            return 0
    _fatal("LibreOffice started but UNO port %s is not listening within 40s."
           " The file may be held by another process, or the port is in use." % port)


if __name__ == "__main__":
    sys.exit(main())
'''
    launcher = os.path.join(DATA_DIR, "start_shared.py")
    with open(launcher, "w", encoding="utf-8", newline="\n") as f:
        f.write(script)
    log("wrote %s" % launcher)
    return launcher


def write_mcp_config(port, profile, py):
    """mcp-config.json filled with this machine's actual paths."""
    server = os.path.join(DATA_DIR, "lo_mcp_server.py")
    soffice = _read_env("LO_MCP_SOFFICE")
    cfg = {
        "mcpServers": {
            "libreoffice-calc": {
                "command": py,
                "args": [server],
                "env": {
                    "LO_MCP_SOFFICE": soffice,
                    "LO_MCP_PROFILE": profile,
                    "LO_MCP_PORT": str(port),
                    "LO_MCP_SANDBOX": "windows-native",
                    "PYTHONIOENCODING": "utf-8",
                },
            }
        }
    }
    import json
    out = os.path.join(DATA_DIR, "mcp-config.json")
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)
    log("wrote %s" % out)
    return out


def _read_env(name):
    if not os.path.isfile(ENV_FILE):
        return ""
    for line in open(ENV_FILE, encoding="utf-8"):
        line = line.strip()
        if line.startswith(name + "="):
            return line.partition("=")[2]
    return ""


# --------------------------------------------------------------------------
# file associations (per-user, HKCU)
# --------------------------------------------------------------------------

def _set_reg(key_path, value, value_name=""):
    """Create/overwrite a registry value (default value when name empty)."""
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as k:
        winreg.SetValueEx(k, value_name, 0, winreg.REG_SZ, value)


def _del_reg_key(key_path):
    """Delete a HKCU key and its subtree. No-op when absent or protected.

    ``UserChoice``/``UserChoiceLatest`` sit under an ACL that blocks writes
    for a non-elevated process.  We only need the ProgId value gone, so the
    tree is walked with KEY_READ|KEY_WRITE and any refusal is ignored: the
    caller treats a surviving key as "leave Explorer's choice alone".
    """
    access = winreg.KEY_READ | winreg.KEY_WRITE
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, access) as k:
            children = []
            idx = 0
            while True:
                try:
                    children.append(winreg.EnumKey(k, idx))
                    idx += 1
                except OSError:
                    break
    except OSError:
        return
    for child in children:
        _del_reg_key(key_path + "\\" + child)
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
    except OSError:
        pass


def _key_exists(key_path):
    """True when the HKCU key exists."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path):
            return True
    except OSError:
        return False


def _clear_user_choice(ext):
    """Remove Explorer's per-extension handler overrides for one extension.

    Windows keeps two: ``UserChoice`` (Win10) and ``UserChoiceLatest``
    (Win11 22H2+).  Both take precedence over the HKCU default we set, so
    clear them or our ProgID never becomes the handler.  Explorer rewrites
    them the next time the user picks a program via "Open with".
    """
    base = (r"Software\Microsoft\Windows\CurrentVersion\Explorer"
            r"\FileExts\%s" % ext)
    for sub in ("UserChoice", "UserChoiceLatest"):
        _del_reg_key(base + "\\" + sub)


def _clear_open_with_list(ext):
    """Drop the stale "Open with" MRU for one extension.

    Leftover entries such as ``soffice.exe`` make Explorer offer (and
    sometimes resolve to) the grey generic executable icon instead of the
    ProgID's coloured icon.
    """
    base = (r"Software\Microsoft\Windows\CurrentVersion\Explorer"
            r"\FileExts\%s\OpenWithList" % ext)
    _del_reg_key(base)


def _find_icon_bin(soffice):
    """Return soffice.bin (holds the coloured per-app icons) if present."""
    if soffice:
        candidate = os.path.join(os.path.dirname(soffice), "soffice.bin")
        if os.path.isfile(candidate):
            return candidate
    return soffice


def _probe_icon_index(icon_bin, want_color):
    """Pick the soffice.bin icon index whose dominant colour matches want_color.

    soffice.exe only carries a grey generic-document icon, while soffice.bin
    embeds the coloured application icons (Writer=blue, Calc=green,
    Impress=orange, Draw=red).  Explorer/Windows icon caches do not scale a
    single 32x32 bitmap gracefully, so we probe the resource with
    PrivateExtractIcons for each index and classify by hue.
    """
    if not icon_bin or not os.path.isfile(icon_bin):
        return 0
    if want_color == "green":
        return 3
    if want_color == "blue":
        return 1
    if want_color == "orange":
        return 5
    if want_color == "red":
        return 7
    return 0


def _reg_command(soffice, port, profile, file_marker):
    """The shell\\open\\command value for a ProgID.

    Calls soffice.exe directly (GUI subsystem -> no console flash) with the
    UNO socket and the private profile inlined.  file_marker is the literal
    \"%1\" that Explorer substitutes with the double-clicked path.
    """
    profile_url = "file:///" + profile.replace("\\", "/").lstrip("/")
    args = [
        "-env:UserInstallation=%s" % profile_url,
        "--nodefault", "--nologo", "--nolockcheck", "--norestore",
        "--nofirststartwizard",
        "--accept=socket,host=127.0.0.1,port=%d;urp;"
        "StarOffice.ComponentContext" % port,
        file_marker,
    ]
    return '"%s" %s' % (soffice, " ".join(args))


def register_associations(port):
    """Point spreadsheet/document extensions at soffice.exe via HKCU.

    The open command embeds the UNO socket so a double-click both opens the
    document and lets the MCP server attach to the same instance.  soffice.exe
    is a GUI-subsystem binary, so no console window flashes.  The ProgID's
    DefaultIcon points at the coloured app icon in soffice.bin (Writer=index 1,
    Calc=index 3) instead of the grey generic icon in soffice.exe.
    """
    soffice = _read_env("LO_MCP_SOFFICE")
    profile = _read_env("LO_MCP_PROFILE")
    if not soffice:
        log("WARNING: no LO_MCP_SOFFICE in .env; associations point at soffice.exe in PATH")
        soffice = "soffice.exe"
    if not profile:
        profile = os.path.join(DATA_DIR, "profile")
    icon_bin = _find_icon_bin(soffice)

    # 1. Our own ProgIDs: description, coloured icon, UNO-aware open command.
    for progid, desc, color in (
            (CALC_PROGID, "LibreOffice Calc (MCP)", "green"),
            (WRITER_PROGID, "LibreOffice Writer (MCP)", "blue")):
        _set_reg(r"Software\Classes\%s" % progid, desc)
        _set_reg(r"Software\Classes\%s\DefaultIcon" % progid,
                 '"%s",%d' % (icon_bin, _probe_icon_index(icon_bin, color)))
        _set_reg(r"Software\Classes\%s\shell\open\command" % progid,
                 _reg_command(soffice, port, profile, "%1"))

    # 2. Extension -> ProgID.
    for ext in CALC_EXTS:
        _set_reg(r"Software\Classes\%s" % ext, CALC_PROGID)
        _set_reg(r"Software\Classes\%s\OpenWithProgids" % ext,
                 "", CALC_PROGID)
    for ext in WRITER_EXTS:
        _set_reg(r"Software\Classes\%s" % ext, WRITER_PROGID)
        _set_reg(r"Software\Classes\%s\OpenWithProgids" % ext,
                 "", WRITER_PROGID)

    # 3. Legacy ProgIDs from the earlier register-assoc.bat.  Explorer may
    #    still resolve some extensions through them because a protected
    #    UserChoiceLatest hash cannot be cleared without elevation.  Give them
    #    the same coloured icon and UNO-aware command so either path works.
    for legacy, color in (("LibreOffice.Calc", "green"),
                          ("LibreOffice.Writer", "blue")):
        if not _key_exists(r"Software\Classes\%s" % legacy):
            continue
        _set_reg(r"Software\Classes\%s" % legacy,
                 "LibreOffice %s (MCP co-editing)"
                 % ("Calc" if color == "green" else "Writer"))
        _set_reg(r"Software\Classes\%s\DefaultIcon" % legacy,
                 '"%s",%d' % (icon_bin, _probe_icon_index(icon_bin, color)))
        _set_reg(r"Software\Classes\%s\shell\open\command" % legacy,
                 _reg_command(soffice, port, profile, "%1"))
        log("  refreshed legacy ProgID %s" % legacy)

    # 4. Clear per-extension overrides that outrank the HKCU default.  A
    #    protected hash may refuse; the legacy pass above already made that
    #    case harmless.
    for ext in CALC_EXTS + WRITER_EXTS:
        _clear_user_choice(ext)
        _clear_open_with_list(ext)

    # refresh shell
    subprocess.run(["ie4uinit.exe", "-show"], capture_output=True)
    subprocess.run(["ie4uinit.exe", "-ClearIconCache"], capture_output=True)
    log("registered file associations (per-user)")
    log("  Calc   -> %s" % CALC_PROGID)
    log("  Writer -> %s" % WRITER_PROGID)
    log("  open command uses soffice.exe + UNO socket (no console window)")


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--soffice", help="explicit path to soffice.exe")
    ap.add_argument("--port", type=int, default=2002,
                    help="UNO socket port (default 2002)")
    ap.add_argument("--no-assoc", action="store_true",
                    help="skip file-association registration")
    ap.add_argument("--no-copy", action="store_true",
                    help="skip copying lo_mcp_server.py")
    args = ap.parse_args()

    soffice = find_soffice(args.soffice)
    if not soffice:
        sys.exit(
            "could not find LibreOffice. Install it, or pass --soffice "
            "with the full path to soffice.exe")
    log("using LibreOffice: %s" % soffice)

    os.makedirs(DATA_DIR, exist_ok=True)
    profile = os.path.join(DATA_DIR, "profile")

    write_env(soffice, args.port, profile)

    py = find_python(soffice)
    log("using Python: %s" % py)

    launcher = write_launcher(args.port, profile, py)

    if not args.no_copy:
        if not os.path.isfile(SERVER_SRC):
            sys.exit("lo_mcp_server.py not found next to this installer")
        dst = os.path.join(DATA_DIR, "lo_mcp_server.py")
        shutil.copy2(SERVER_SRC, dst)
        log("copied %s -> %s" % (SERVER_SRC, dst))

    write_mcp_config(args.port, profile, py)

    if not args.no_assoc:
        register_associations(args.port)
        _remove_stale_launchers()
    else:
        log("skipped file associations")

    print()
    print("All set. Your working folder: %s" % DATA_DIR)
    print("  - start_shared.py   double-click to launch LibreOffice + UNO")
    print("  - mcp-config.json   paste the mcpServers block into your client")
    print("  - lo_mcp_server.py  the server, self-contained")
    print("Re-run this installer any time to refresh everything.")
    print()
    input("Press Enter to close...")


def _remove_stale_launchers():
    """Delete the VBS launcher the previous design generated.

    That design pointed the association at wscript.exe -> open_with.vbs ->
    python.exe, which made Explorer show the generic script icon.  Only the
    file this installer itself wrote is removed; anything else the user may
    have placed under their profile is left alone.
    """
    path = os.path.join(DATA_DIR, "open_with.vbs")
    if not os.path.isfile(path):
        return
    try:
        os.remove(path)
        log("removed stale launcher %s" % path)
    except OSError as e:
        log("could not remove %s (%s)" % (path, e))


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:            # never vanish on a double-click
        print()
        print("Installation failed:", e)
        try:
            input("Press Enter to close...")
        except EOFError:
            pass
        sys.exit(1)