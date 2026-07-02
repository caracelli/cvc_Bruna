# Abre o Chrome num PERFIL DEDICADO com a porta de depuração ligada.
#
# Por que perfil dedicado:
#   - Desde o Chrome 136, a porta de depuração NAO funciona no perfil padrao.
#   - Assim a automacao nao mexe no seu Chrome do dia a dia.
#   - O login no Outlook (com MFA) fica salvo neste perfil: voce loga 1 vez so.
#
# COMO USAR:
#   1. Rode este script (clique direito > Executar com PowerShell, ou no terminal).
#   2. Na janela do Chrome que abrir, acesse https://outlook.office.com e logue.
#   3. Pode deixar essa janela aberta. Depois rode:  python test_navegador.py

$chrome  = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$perfil  = "C:\Users\user\chrome_cvc_profile"   # fora do OneDrive (evita sync/lock)
$porta   = 9222

if (-not (Test-Path $perfil)) {
    New-Item -ItemType Directory -Force -Path $perfil | Out-Null
}

Write-Host "Abrindo Chrome (perfil CVC) na porta de depuracao $porta..." -ForegroundColor Cyan
Write-Host "Perfil: $perfil" -ForegroundColor DarkGray

& $chrome `
    "--remote-debugging-port=$porta" `
    "--user-data-dir=$perfil" `
    "--no-first-run" `
    "--no-default-browser-check" `
    "https://outlook.office.com/mail/"

Write-Host ""
Write-Host "Pronto. Se for a 1a vez, FACA LOGIN no Outlook nessa janela." -ForegroundColor Green
Write-Host "Depois rode:  python test_navegador.py" -ForegroundColor Green
