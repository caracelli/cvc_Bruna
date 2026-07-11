<#
================================================================================
  build.ps1 - Gera o pacote .EXE do CVC-Trata-Forms (PyInstaller onedir)
================================================================================
  Reproduz o bundle 'CVC-Trata-Forms.zip' do zero:
    1. garante o venv (.venv) com as deps + PyInstaller
    2. roda o PyInstaller (onedir, windowed) empacotando o driver do Playwright
    3. injeta os assets do bundle (COMO_USAR.txt + dados/config.xml)
    4. compacta o zip final (com barras '/', padrao do formato)

  USO:
    powershell -ExecutionPolicy Bypass -File build.ps1
    powershell -ExecutionPolicy Bypass -File build.ps1 -Recreatevenv   # recria o venv

  REQUISITO: Python 3.10+ instalado (procura no PATH e no local padrao do winget).
  O pacote gerado NAO precisa de Python na maquina do cliente.
================================================================================
#>
param(
    [switch]$RecreateVenv
)

$ErrorActionPreference = "Stop"
$Repo   = $PSScriptRoot
$Venv   = Join-Path $Repo ".venv"
$VPy    = Join-Path $Venv "Scripts\python.exe"
$Dist   = Join-Path $Repo "dist\CVC-Trata-Forms"
$Pack   = Join-Path $Repo "packaging"
$ZipOut = Join-Path $Repo "CVC-Trata-Forms.zip"

function Find-Python {
    foreach ($c in @("python", "py")) {
        $g = Get-Command $c -ErrorAction SilentlyContinue
        if ($g -and $g.Source -notmatch "WindowsApps") { return $g.Source }
    }
    $guess = Get-ChildItem "$env:LOCALAPPDATA\Programs\Python" -Filter python.exe -Recurse -ErrorAction SilentlyContinue |
             Select-Object -First 1
    if ($guess) { return $guess.FullName }
    throw "Python 3.10+ nao encontrado. Instale com: winget install Python.Python.3.14"
}

# --- 1. venv -----------------------------------------------------------------
if ($RecreateVenv -and (Test-Path $Venv)) { Remove-Item $Venv -Recurse -Force }
if (-not (Test-Path $VPy)) {
    $py = Find-Python
    Write-Host ">> Criando venv com: $py"
    & $py -m venv $Venv
    & $VPy -m pip install --upgrade pip
    & $VPy -m pip install -r (Join-Path $Repo "requirements.txt") pyinstaller
} else {
    Write-Host ">> venv ja existe (use -RecreateVenv para recriar)."
}

# --- 2. PyInstaller ----------------------------------------------------------
Write-Host ">> Rodando PyInstaller (onedir, windowed)..."
Push-Location $Repo
try {
    & $VPy -m PyInstaller --noconfirm --clean --onedir --windowed `
        --name CVC-Trata-Forms --paths src `
        --collect-all playwright --collect-all pystray --collect-all keyring `
        --collect-all uiautomation --collect-all PIL `
        app.py
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller falhou (exit $LASTEXITCODE)." }
} finally {
    Pop-Location
}

# --- 3. Assets do bundle -----------------------------------------------------
Write-Host ">> Injetando assets (COMO_USAR.txt + dados/config.xml)..."
Copy-Item (Join-Path $Pack "COMO_USAR.txt") (Join-Path $Dist "COMO_USAR.txt") -Force
New-Item -ItemType Directory -Force -Path (Join-Path $Dist "dados") | Out-Null
$cfg = Join-Path $Pack "config.xml"
if (Test-Path $cfg) {
    Copy-Item $cfg (Join-Path $Dist "dados\config.xml") -Force
} else {
    Write-Warning "packaging\config.xml nao existe; usando config.example.xml (dry_run, sem senha)."
    Copy-Item (Join-Path $Repo "config.example.xml") (Join-Path $Dist "dados\config.xml") -Force
}

# --- 4. Zip final (barras '/', via Python) -----------------------------------
Write-Host ">> Validando o exe (--check)..."
$p = Start-Process -FilePath (Join-Path $Dist "CVC-Trata-Forms.exe") -ArgumentList "--check" -Wait -PassThru
if ($p.ExitCode -ne 0) { throw "Validacao --check falhou (exit $($p.ExitCode)): alguma lib nao empacotou." }
Write-Host "   OK: libs pesadas empacotadas."

if (Test-Path $ZipOut) {
    Move-Item $ZipOut (Join-Path $Repo "CVC-Trata-Forms.OLD.zip") -Force
    Write-Host ">> Zip anterior -> CVC-Trata-Forms.OLD.zip"
}
Write-Host ">> Compactando o pacote..."
$base = Join-Path $Repo "CVC-Trata-Forms"
& $VPy -c "import shutil; shutil.make_archive(r'$base', 'zip', root_dir=r'$Repo\dist', base_dir='CVC-Trata-Forms')"

$mb = [math]::Round((Get-Item $ZipOut).Length / 1MB, 1)
Write-Host ""
Write-Host "================================================================"
Write-Host "  PRONTO: $ZipOut  ($mb MB)"
Write-Host "================================================================"
