#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Start LibreOffice with a UNO socket listening, so the MCP server can attach
to it and co-edit documents you already have open.

Windows helper; on Linux/macOS you can start LibreOffice normally and add
"--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext".

Usage:
    python scripts/start_shared.py               # open a blank Calc document
    python scripts/start_shared.py path/file.ods # open a specific file
"""
import os
import subprocess
import sys
import time

LO_HOME = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PORT = os.environ.get("LO_MCP_PORT", "2002")
PROFILE = os.environ.get("LO_MCP_PROFILE",
                         os.path.join(LO_HOME, "data", "profile-user"))


def find_soffice():
    env = os.environ.get("LO_MCP_SOFFICE", "")
    if env and os.path.isfile(env):
        return env
    for name in ("soffice", "libreoffice"):
        found = subprocess.run(
            ["where", name] if os.name == "nt" else ["which", name],
            capture_output=True, text=True)
        if found.returncode == 0:
            return found.stdout.splitlines()[0].strip()
    raise RuntimeError("soffice not found; set LO_MCP_SOFFICE")


def port_listening():
    if os.name == "nt":
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "if (Get-NetTCPConnection -LocalPort %s -State Listen "
             "-ErrorAction SilentlyContinue) { 'yes' } else { 'no' }" % PORT],
            capture_output=True, text=True)
        return r.stdout.strip() == "yes"
    r = subprocess.run(
        ["sh", "-c", "ss -ltn 2>/dev/null | grep -q ':%s ' || "
         "netstat -ltn 2>/dev/null | grep -q ':%s '" % (PORT, PORT)],
        capture_output=True, text=True)
    return r.returncode == 0


def main():
    os.makedirs(PROFILE, exist_ok=True)
    profile_url = "file:///" + PROFILE.replace("\\", "/").lstrip("/")
    soffice = find_soffice()

    cmd = [
        soffice,
        "-env:UserInstallation=%s" % profile_url,
        "--nodefault", "--nologo", "--nolockcheck", "--norestore",
        "--nofirststartwizard",
        "--accept=socket,host=127.0.0.1,port=%s;urp;StarOffice.ComponentContext"
        % PORT,
    ]
    if len(sys.argv) > 1:
        target = os.path.abspath(sys.argv[1])
        if not os.path.isfile(target):
            sys.exit("no such file: %s" % target)
        cmd.append(target)
    else:
        cmd.append("--calc")

    # Detach from this process's tree so LibreOffice survives the helper
    # exiting. Win32_Process.Create does that on Windows; setsid on POSIX.
    if os.name == "nt":
        quoted = " ".join('"%s"' % c if " " in c or ";" in c else c
                          for c in cmd)
        ps = ("$r = Invoke-CimMethod -ClassName Win32_Process "
              "-MethodName Create -Arguments @{ CommandLine = '%s' }; "
              "Write-Output $r.ReturnValue" % quoted.replace("'", "''"))
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                           capture_output=True, text=True)
        if r.stdout.strip() != "0":
            sys.exit("failed to start LibreOffice: %s" % r.stderr[:300])
    else:
        subprocess.Popen(cmd, start_new_session=True,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    for _ in range(40):
        time.sleep(1)
        if port_listening():
            print("LibreOffice up, UNO port %s ready" % PORT)
            return 0
    print("LibreOffice started but port %s not listening within 40s" % PORT)
    return 0


if __name__ == "__main__":
    sys.exit(main())