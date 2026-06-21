<#
build_installer.ps1 — build a (optionally signed) Windows installer for JewishGenealogySearch.

WHAT IT DOES, in order:
  1. find the project's virtual-env Python
  2. generate the app icon (config/app_icon.ico) and a Windows version resource
  3. run PyInstaller (packaging/JewishGenealogySearch.spec) → dist/JewishGenealogySearch/
  4. copy Playwright's Chromium next to the .exe (so the app is self-contained)
  5. Authenticode-sign the .exe            (only if you supply a code-signing cert)
  6. compile the installer with Inno Setup → Output/JewishGenealogySearch-Setup-<ver>.exe
  7. Authenticode-sign the installer        (only if you supply a code-signing cert)

YOU RUN IT yourself after any code change, simply:
    .\build_installer.ps1                       # unsigned test build
    .\build_installer.ps1 -Version 1.2.0        # set the version
    $env:CODESIGN_PFX = "C:\path\cert.pfx"; $env:CODESIGN_PFX_PASSWORD = "secret"
    .\build_installer.ps1 -Version 1.2.0        # signed build (cert from env)

ONE-TIME PREREQUISITES (see packaging/BUILD.md for details):
  • pip install pyinstaller pillow           (into the project venv)
  • Inno Setup 6           https://jrsoftware.org/isinfo.php   (for the installer)
  • a code-signing certificate (.pfx) from a CA   (for signing — optional but recommended;
    signing is what gives anti-tamper integrity and avoids antivirus / SmartScreen warnings)
  • Windows SDK signtool.exe                 (ships with Visual Studio / Windows SDK)

NOTES ON "protection":
  • Authenticode signing = tamper-evidence: any edit to the .exe breaks the signature, and
    a signed+timestamped binary builds SmartScreen reputation and avoids AV false positives.
  • UPX compression is intentionally OFF in the spec (UPX trips antivirus heuristics).
  • Python apps can still be decompiled; for stronger obfuscation install PyArmor and see
    packaging/BUILD.md. No Python packaging is truly crack-proof — signing is the realistic,
    industry-standard protection here.
#>
param(
    [string]$Version       = "1.0.0",
    [string]$CertPfx       = $env:CODESIGN_PFX,
    [string]$CertPass      = $env:CODESIGN_PFX_PASSWORD,
    [string]$TimestampUrl  = "http://timestamp.digicert.com",
    [switch]$SkipInstaller,
    [switch]$Clean
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
$Root = $PSScriptRoot

function Info($m)  { Write-Host "==> $m" -ForegroundColor Cyan }
function Ok($m)    { Write-Host "    $m" -ForegroundColor Green }
function Warn($m)  { Write-Host "    ! $m" -ForegroundColor Yellow }

# ── 1. find the venv python ──────────────────────────────────────────────────
$py = $null
foreach ($c in @(".\.venv\Scripts\python.exe", ".\venv\Scripts\python.exe")) {
    if (Test-Path $c) { $py = (Resolve-Path $c).Path; break }
}
if (-not $py) {
    if (Get-Command python -ErrorAction SilentlyContinue) { $py = "python"; Warn "no venv found — using system python" }
    else { throw "Python not found. Create the venv and 'pip install -r requirements.txt' first." }
}
Info "Python: $py"

# verify build tooling is importable
& $py -c "import PyInstaller, PIL" 2>$null
if ($LASTEXITCODE -ne 0) { throw "PyInstaller / Pillow missing in the venv. Run:  $py -m pip install pyinstaller pillow" }

# ── 2. icon + version resource ───────────────────────────────────────────────
Info "Generating config/app_icon.ico"
& $py -c "from PIL import Image; Image.open('config/app_icon.png').save('config/app_icon.ico', sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])"
Ok "icon written"

Info "Writing packaging/version_info.txt (version $Version)"
$v = ($Version.Split('.') + @('0','0','0','0'))[0..3] -join ','
@"
VSVersionInfo(
  ffi=FixedFileInfo(filevers=($v), prodvers=($v), mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0,0)),
  kids=[
    StringFileInfo([StringTable(u'040904B0', [
      StringStruct(u'CompanyName', u'Alla Khananashvili'),
      StringStruct(u'FileDescription', u'Jewish Genealogy Search — genealogy meta-search'),
      StringStruct(u'FileVersion', u'$Version'),
      StringStruct(u'InternalName', u'JewishGenealogySearch'),
      StringStruct(u'OriginalFilename', u'JewishGenealogySearch.exe'),
      StringStruct(u'ProductName', u'Jewish Genealogy Search'),
      StringStruct(u'ProductVersion', u'$Version')])]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
"@ | Out-File -FilePath "packaging\version_info.txt" -Encoding utf8
Ok "version resource written"

# ── 3. PyInstaller ───────────────────────────────────────────────────────────
if ($Clean -and (Test-Path "build")) { Remove-Item "build" -Recurse -Force }
if (Test-Path "dist\JewishGenealogySearch") { Remove-Item "dist\JewishGenealogySearch" -Recurse -Force }
Info "Running PyInstaller (this can take a few minutes)…"
& $py -m PyInstaller --noconfirm --clean "packaging\JewishGenealogySearch.spec"
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed." }
$DistApp = Join-Path $Root "dist\JewishGenealogySearch"
if (-not (Test-Path (Join-Path $DistApp "JewishGenealogySearch.exe"))) { throw "Build produced no JewishGenealogySearch.exe" }
Ok "built dist\JewishGenealogySearch\JewishGenealogySearch.exe"

# ── 4. bundle the ACTIVE Playwright Chromium next to the exe ──────────────────
# Pack the Chromium that Playwright ACTUALLY uses (the user already has it — don't
# download another; her older/incomplete chromium-* folders must NOT be shipped).
Info "Bundling the Chromium Playwright actually uses"
$pwCache = Join-Path $env:LOCALAPPDATA "ms-playwright"
$activeExe = & $py -c "from playwright.sync_api import sync_playwright; p=sync_playwright().start(); print(p.chromium.executable_path); p.stop()" 2>$null |
             Select-Object -Last 1
$activeDir = $null
if ($activeExe -and (Test-Path $activeExe)) {
    # …\ms-playwright\chromium-1194\chrome-win\chrome.exe → …\chromium-1194
    $activeDir = Split-Path (Split-Path $activeExe -Parent) -Parent
    Ok "active Chromium: $(Split-Path $activeDir -Leaf)"
} elseif (-not (Test-Path $pwCache)) {
    Warn "no Playwright Chromium found — installing it now…"
    & $py -m playwright install chromium
}
$dest = Join-Path $DistApp "ms-playwright"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
$copied = 0
foreach ($d in Get-ChildItem $pwCache -Directory -ErrorAction SilentlyContinue |
                 Where-Object {
                     ($_.Name -match '^(ffmpeg|winldd)') -or
                     ($activeDir -and ($_.FullName -ieq $activeDir)) -or
                     ((-not $activeDir) -and ($_.Name -match '^chromium'))
                 }) {
    Copy-Item $d.FullName (Join-Path $dest $d.Name) -Recurse -Force
    $copied++
}
if ($copied -eq 0) { Warn "no chromium folder copied — the packaged app may fail to open a browser" }
else { Ok "copied $copied Playwright component(s) (active Chromium only)" }

# ── 5. find signtool + sign helper ───────────────────────────────────────────
$signtool = (Get-Command signtool.exe -ErrorAction SilentlyContinue).Source
if (-not $signtool) {
    $signtool = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe" -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
}
function Sign-File($path) {
    if (-not $CertPfx) { Warn "no code-signing cert (CODESIGN_PFX) — leaving '$([IO.Path]::GetFileName($path))' UNSIGNED"; return }
    if (-not (Test-Path $CertPfx)) { throw "Cert not found: $CertPfx" }
    if (-not $signtool) { throw "signtool.exe not found (install the Windows SDK) — cannot sign." }
    Info "Signing $([IO.Path]::GetFileName($path))"
    & $signtool sign /fd SHA256 /f $CertPfx /p $CertPass /tr $TimestampUrl /td SHA256 $path
    if ($LASTEXITCODE -ne 0) { throw "signing failed for $path" }
    Ok "signed"
}

Sign-File (Join-Path $DistApp "JewishGenealogySearch.exe")

if ($SkipInstaller) { Info "Done (‑SkipInstaller): portable build in dist\JewishGenealogySearch"; return }

# ── 6. Inno Setup installer ──────────────────────────────────────────────────
$iscc = (Get-Command ISCC.exe -ErrorAction SilentlyContinue).Source
if (-not $iscc) {
    foreach ($c in @("${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "${env:ProgramFiles}\Inno Setup 6\ISCC.exe")) {
        if (Test-Path $c) { $iscc = $c; break }
    }
}
if (-not $iscc) { throw "Inno Setup (ISCC.exe) not found. Install from https://jrsoftware.org/isinfo.php (or run with -SkipInstaller for a portable build)." }
Info "Compiling installer with Inno Setup"
& $iscc "/DMyAppVersion=$Version" "packaging\installer.iss"
if ($LASTEXITCODE -ne 0) { throw "Inno Setup compile failed." }
$setup = Join-Path $Root "Output\JewishGenealogySearch-Setup-$Version.exe"
if (-not (Test-Path $setup)) { throw "Installer not produced." }
Ok "built $setup"

# ── 7. sign the installer ────────────────────────────────────────────────────
Sign-File $setup

Write-Host ""
Info "BUILD COMPLETE"
Ok   "Installer : $setup"
Ok   "Portable  : $DistApp"
if (-not $CertPfx) { Warn "Build is UNSIGNED — set CODESIGN_PFX / CODESIGN_PFX_PASSWORD for a signed, antivirus-friendly release." }
