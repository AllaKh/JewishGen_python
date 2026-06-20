# Building the JewishGen Search installer

You build the Windows installer yourself by running, from the project root:

```powershell
.\build_installer.ps1
```

That produces:

* `dist\JewishGenSearch\` — the portable app folder (the `.exe` + everything it needs), and
* `Output\JewishGenSearch-Setup-<version>.exe` — the installer to hand to other people.

Run it again any time after you change the code — it rebuilds from scratch.

```powershell
.\build_installer.ps1 -Version 1.3.0      # set the version number
.\build_installer.ps1 -SkipInstaller      # only the portable folder, no installer
.\build_installer.ps1 -Clean              # also wipe PyInstaller's build cache
```

---

## One-time setup

1. **Build tools (into the project venv):**
   ```powershell
   .\venv\Scripts\python -m pip install pyinstaller pillow
   ```
2. **Playwright browser** (the app drives a real Chromium):
   ```powershell
   .\venv\Scripts\python -m playwright install chromium
   ```
   The build script copies this Chromium next to the `.exe`, so the installed app is
   self-contained (the end user does **not** need Python or Playwright).
3. **Inno Setup 6** — https://jrsoftware.org/isinfo.php (this is what makes the installer).
4. **Windows SDK** (for `signtool.exe`) — only needed if you sign (see below). It comes with
   Visual Studio or the standalone Windows SDK.

---

## Signing (anti-tamper + antivirus / SmartScreen)

Signing is the realistic, industry-standard protection for a Windows app:

* It makes the binary **tamper-evident** — any edit to the `.exe` invalidates the signature.
* A signed **and timestamped** binary builds **SmartScreen reputation** and stops most
  **antivirus false positives** and the "unknown publisher" warning.

To sign, give the script your **code-signing certificate** (a `.pfx` you buy from a CA such
as DigiCert, Sectigo, etc.):

```powershell
$env:CODESIGN_PFX = "C:\path\to\your-codesign.pfx"
$env:CODESIGN_PFX_PASSWORD = "your-pfx-password"
.\build_installer.ps1 -Version 1.3.0
```

Both the `.exe` and the installer get signed. Without a cert the build still works, but it
is **unsigned** (fine for testing on your own machine).

### Just testing? make a self-signed cert
```powershell
$c = New-SelfSignedCertificate -Type CodeSigning -Subject "CN=JewishGen Search Test" -CertStoreLocation Cert:\CurrentUser\My
$pw = ConvertTo-SecureString "test123" -AsPlainText -Force
Export-PfxCertificate -Cert $c -FilePath ".\codesign-test.pfx" -Password $pw
$env:CODESIGN_PFX = ".\codesign-test.pfx"; $env:CODESIGN_PFX_PASSWORD = "test123"
```
A self-signed cert verifies the signing pipeline but is **not** trusted by other machines —
for real distribution you need a CA-issued certificate.

---

## Stronger anti-crack (optional)

Python bytecode can be decompiled. PyInstaller already ships bytecode (not source) and UPX
is disabled (UPX trips antivirus). For stronger obfuscation install **PyArmor**
(https://pyarmor.dashingsoft.com/) and obfuscate before packaging. Be aware that **no**
Python packaging is truly crack-proof — Authenticode signing (integrity + reputation) is
the protection that actually matters for distribution.

---

## What is generated (and git-ignored)

`build/`, `dist/`, `Output/`, `config/app_icon.ico`, `packaging/version_info.txt` and any
`*.pfx` are build artifacts — they are not committed.
