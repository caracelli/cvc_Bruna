"""
Credenciais para o login automatico (Outlook).

Duas origens, nesta ordem de prioridade:
  1. Arquivo local  credenciais.xml  (portavel para a maquina do cliente)
  2. Gerenciador de Credenciais do Windows (keyring/DPAPI)

API:
    obter_outlook() -> (email, senha)  (ou (None, None) se nao houver)
    salvar_outlook(email, senha)       (grava no cofre do Windows)
    apagar_outlook()
"""

import os
import xml.etree.ElementTree as ET

import keyring

from cvc_acessos.infrastructure.sistema.caminhos import RAIZ

SERVICO = "cvc-gestao-acessos"
CHAVE_USER = "outlook::user"   # guarda QUAL e-mail usar
ARQ_XML = os.path.join(RAIZ, "credenciais.xml")


def _ler_xml():
    """Le email/senha do credenciais.xml, se existir."""
    if not os.path.exists(ARQ_XML):
        return None, None
    try:
        raiz = ET.parse(ARQ_XML).getroot()
        ol = raiz.find("outlook")
        if ol is None:
            return None, None
        email = (ol.findtext("email") or "").strip() or None
        senha = (ol.findtext("senha") or "").strip() or None
        return email, senha
    except Exception:
        return None, None


def salvar_outlook(email, senha):
    keyring.set_password(SERVICO, CHAVE_USER, email)
    keyring.set_password(SERVICO, email, senha)


def obter_outlook():
    # 1) config.xml centralizado (prioridade)
    try:
        from cvc_acessos.infrastructure.config.config_app import OUTLOOK_EMAIL, OUTLOOK_SENHA
        if OUTLOOK_EMAIL and OUTLOOK_SENHA:
            return OUTLOOK_EMAIL, OUTLOOK_SENHA
    except Exception:
        pass
    # 2) credenciais.xml (compatibilidade)
    email, senha = _ler_xml()
    if email and senha:
        return email, senha
    # 3) cofre do Windows como alternativa
    email = keyring.get_password(SERVICO, CHAVE_USER)
    if not email:
        return None, None
    senha = keyring.get_password(SERVICO, email)
    return email, senha


def apagar_outlook():
    email = keyring.get_password(SERVICO, CHAVE_USER)
    try:
        if email:
            keyring.delete_password(SERVICO, email)
        keyring.delete_password(SERVICO, CHAVE_USER)
    except keyring.errors.PasswordDeleteError:
        pass
