"""
Camada de DOMINIO — regras puras (sem I/O, sem navegador).
Fáceis de testar isoladamente (rodam no CI).
"""

# prefixos do assunto Forms a remover p/ obter o titulo limpo (multi-idioma)
PREFIXOS = ["Nova resposta de ", "Nova resposta a ",
            "Nueva respuesta de ", "Nueva respuesta a ",
            "New response for ", "New response to "]


def limpar_titulo(raw):
    """Do texto bruto do cabecalho do e-mail Forms -> titulo limpo:
    corta 'Resumir'/'Summarize'/quebra de linha e remove o prefixo
    'Nova resposta de ...' (multi-idioma)."""
    raw = (raw or "").strip()
    for corte in ["Resumir", "Summarize", "\n"]:
        raw = raw.split(corte)[0].strip()
    for pre in PREFIXOS:
        if raw.lower().startswith(pre.lower()):
            raw = raw[len(pre):].strip()
            break
    return raw


def tem_cat_jira(categorias, prefixo):
    """True se alguma categoria comeca com o prefixo (ex.: 'JIRA -')."""
    pref = (prefixo or "").lower()
    return any((c or "").lower().startswith(pref) for c in (categorias or []))


def casa_remetente(resumo, termo):
    """True se o resumo do e-mail contem o termo do remetente (ou termo vazio)."""
    if not termo:
        return True
    return termo.lower() in (resumo or "").lower()
