"""
Captura de diagnostico VISUAL (print + HTML + txt) da pagina atual.

Usado quando algo falha na maquina do cliente e o log em texto nao mostra a
tela (menu que nao abriu, botao que nao respondeu, banner diferente...).
Generico: recebe qualquer 'page' do Playwright (Jira OU Outlook).
"""

import os
import time

from cvc_trata_forms.infrastructure.sistema.caminhos import RAIZ


def salvar_diagnostico(page, etapa, log=print):
    """Salva print + HTML + txt da pagina ATUAL em <RAIZ>/diagnostico para
    inspecao remota. Nunca lanca excecao. Retorna o caminho-base ou None."""
    try:
        pasta = os.path.join(RAIZ, "diagnostico")
        os.makedirs(pasta, exist_ok=True)
        carimbo = time.strftime("%Y%m%d_%H%M%S")
        base = os.path.join(pasta, f"{etapa}_{carimbo}")
        try:
            page.screenshot(path=base + ".png", full_page=True)
        except Exception:
            try:
                page.screenshot(path=base + ".png")   # viewport (fallback)
            except Exception as e:
                log(f"   [diag] print falhou: {e}")
        try:
            with open(base + ".html", "w", encoding="utf-8") as f:
                f.write(page.content())
        except Exception as e:
            log(f"   [diag] html falhou: {e}")
        try:
            with open(base + ".txt", "w", encoding="utf-8") as f:
                f.write(f"etapa: {etapa}\n")
                f.write(f"url: {getattr(page, 'url', '?')}\n")
                f.write(f"quando: {carimbo}\n")
        except Exception:
            pass
        log(f"   [diag] diagnostico salvo: {base}.(png|html|txt)")
        return base
    except Exception as e:
        log(f"   [diag] nao consegui salvar diagnostico: {e}")
        return None
