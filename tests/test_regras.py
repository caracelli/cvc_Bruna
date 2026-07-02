"""Testes da camada de DOMINIO (puros, rodam no CI sem navegador)."""
from cvc_acessos.domain.regras import (
    limpar_titulo, tem_cat_jira, casa_remetente,
)


def test_limpar_titulo_remove_prefixo():
    assert limpar_titulo("Nova resposta de Solicitud: Alta") == "Solicitud: Alta"
    assert limpar_titulo("New response for Gestion de Usuarios") == \
        "Gestion de Usuarios"


def test_limpar_titulo_corta_resumir_e_quebra():
    assert limpar_titulo("Nova resposta de X Resumir bla bla") == "X"
    assert limpar_titulo("Nova resposta de Y\nlinha 2") == "Y"


def test_limpar_titulo_vazio():
    assert limpar_titulo("") == ""
    assert limpar_titulo(None) == ""


def test_tem_cat_jira():
    assert tem_cat_jira(["JIRA - GAAR-1"], "JIRA -") is True
    assert tem_cat_jira(["FEITO", "Pendente"], "JIRA -") is False
    assert tem_cat_jira([], "JIRA -") is False
    assert tem_cat_jira(None, "JIRA -") is False


def test_casa_remetente():
    assert casa_remetente("Microsoft Forms Nova resposta", "Microsoft Forms")
    assert casa_remetente("qualquer coisa", "")          # termo vazio = tudo
    assert not casa_remetente("Outro Remetente", "Microsoft Forms")
