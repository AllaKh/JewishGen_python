"""
machine_lock.py — optional HARDWARE BINDING (node-lock) for the packaged app.

A LOCKED package runs ONLY on the machine(s) whose id is baked into
config/machine_lock.txt at build time (written by build_installer.ps1 -MachineId /
build_locked_installer.ps1). This is the offline "one machine only" option — no server,
no payments — at the cost of a one-time manual exchange of the recipient's machine id.

Enforcement rules (fail-open by design, so a misconfiguration never bricks the app):
  * DEV run (not frozen)                       → NEVER enforced (you can always run).
  * Packaged build WITHOUT a lock file         → runs everywhere (an ordinary build).
  * Packaged build WITH config/machine_lock.txt → runs only if THIS machine's id is listed.

The machine id is the Windows MachineGuid (stable per OS install, readable without admin) —
the SAME value get_machine_id.bat prints, so the id the recipient sends matches what the
app reads here. Comparison is case-insensitive and brace-insensitive.
"""
import sys
from pathlib import Path


def _config_dir() -> Path:
    # packaged: config/ sits next to the .exe; dev: the project's config/
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "config"
    return Path(__file__).resolve().parent / "config"


def _norm(s: str) -> str:
    return (s or "").strip().strip("{}").lower()


def machine_id() -> str:
    """This machine's Windows MachineGuid (normalised). Empty string on failure."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Cryptography", 0,
                            winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as k:
            val, _ = winreg.QueryValueEx(k, "MachineGuid")
        return _norm(str(val))
    except Exception:
        return ""


def allowed_ids() -> list:
    """Machine ids this package is locked to (config/machine_lock.txt, one per line;
    blank lines and lines starting with '#' are ignored). Empty list = NOT locked."""
    f = _config_dir() / "machine_lock.txt"
    try:
        if not f.is_file():
            return []
        out = []
        for ln in f.read_text(encoding="utf-8").splitlines():
            n = _norm(ln)
            if n and not ln.lstrip().startswith("#"):
                out.append(n)
        return out
    except Exception:
        return []


def check_or_exit():
    """Enforce the hardware lock. Call ONCE at startup, right after QApplication() exists
    (so a message box can be shown). Never blocks a dev run or an unlocked build."""
    if not getattr(sys, "frozen", False):
        return                                   # dev — never locked
    allowed = allowed_ids()
    if not allowed:                              # ordinary (unlocked) build
        return
    if machine_id() in allowed:                  # this machine is licensed
        return
    try:
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.critical(
            None, "Лицензия",
            "Эта копия программы привязана к другому компьютеру и на этой машине "
            "работать не будет.\n\nОбратитесь к поставщику за версией для вашего "
            "компьютера (понадобится ваш machine id — запустите get_machine_id.bat).")
    except Exception:
        pass
    sys.exit(1)
