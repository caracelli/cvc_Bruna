"""
Cadastro UNICA VEZ das credenciais do Outlook no cofre do Windows.

VOCE roda este script no seu terminal. A senha e digitada SEM aparecer na
tela e guardada criptografada (DPAPI). Ninguem (nem eu) ve a senha.

USO:
    python salvar_credenciais.py
    python salvar_credenciais.py --apagar     # remove as credenciais
"""

import getpass
import sys

from cvc_trata_forms.infrastructure.credenciais.credenciais import salvar_outlook, obter_outlook, apagar_outlook


def main():
    if "--apagar" in sys.argv:
        apagar_outlook()
        print("Credenciais removidas do cofre.")
        return

    email_atual, _ = obter_outlook()
    if email_atual:
        print(f"Ja existe credencial salva para: {email_atual}")
        resp = input("Sobrescrever? (s/N): ").strip().lower()
        if resp != "s":
            print("Mantido. Nada alterado.")
            return

    email = input("E-mail do Outlook (ex: nelsondiniz@ext.cvccorp.com.br): ").strip()
    if not email:
        print("E-mail vazio. Abortado.")
        return
    senha = getpass.getpass("Senha (nao aparece enquanto digita): ")
    if not senha:
        print("Senha vazia. Abortado.")
        return

    salvar_outlook(email, senha)
    print(f"\nOK! Credenciais de {email} guardadas no Gerenciador de "
          "Credenciais do Windows (criptografadas).")
    print("Agora o login automatico pode preencher e-mail+senha sozinho "
          "(o MFA ainda sera com voce).")


if __name__ == "__main__":
    main()
