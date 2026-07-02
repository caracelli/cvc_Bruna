# Abre o EDGE num PERFIL DEDICADO com a porta de depuracao ligada.
#
# Edge usa o mesmo motor do Chrome (Chromium), entao funciona igual.
# Use este se no cliente CVC voce usa o Edge.
#
# Por que perfil dedicado:
#   - Desde o Chromium 136, a porta de depuracao NAO funciona no perfil padrao.
#   - Assim a automacao nao mexe no seu Edge do dia a dia.
#   - O login no Outlook (com MFA) fica salvo neste perfil: voce loga 1 vez so.
#
# COMO USAR:
#   1. Rode este script.
#   2. Na janela do Edge que abrir, acesse https://outlook.office.com e logue.
#   3. Pode deixar aberta. Depois rode:  python test_navegador.py

$edge    = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$perfil  = "C:\Users\user\edge_cvc_profile"   # fora do OneDrive (evita sync/lock)
$porta   = 9222

if (-not (Test-Path $perfil)) {
    New-Item -ItemType Directory -Force -Path $perfil | Out-Null
}

Write-Host "Abrindo Edge (perfil CVC) na porta de depuracao $porta..." -ForegroundColor Cyan
Write-Host "Perfil: $perfil" -ForegroundColor DarkGray

& $edge `
    "--remote-debugging-port=$porta" `
    "--user-data-dir=$perfil" `
    "--no-first-run" `
    "--no-default-browser-check" `
    "https://outlook.office.com/mail/"

Write-Host ""
Write-Host "Pronto. Se for a 1a vez, FACA LOGIN no Outlook nessa janela." -ForegroundColor Green
Write-Host "Depois rode:  python test_navegador.py" -ForegroundColor Green
