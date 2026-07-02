"""
Formulario de configuracao (tkinter) para o config.xml.
- Nao precisa instalar nada (tkinter vem no Python).
- Le o config.xml atual, permite editar e salva de volta.
- Senha mascarada (com opcao 'mostrar').
- Botao 'Abrir config.xml' abre no editor padrao.

Uso: python config_form.py
"""
import os
import sys
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

import tkinter as tk
from tkinter import ttk, messagebox

ARQ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.xml")


# -------------------- leitura dos valores atuais --------------------
def _ler():
    v = {}
    root = None
    if os.path.exists(ARQ):
        try:
            root = ET.parse(ARQ).getroot()
        except Exception:
            root = None

    def g(path, default=""):
        if root is None:
            return default
        el = root.find(path)
        if el is not None and el.text is not None:
            return el.text.strip()
        return default

    v["outlook_email"] = g("outlook/email")
    v["outlook_senha"] = g("outlook/senha")
    v["outlook_url"] = g("outlook/url", "https://outlook.office.com/mail/")
    v["outlook_caixa"] = g("outlook/caixa", "Gestão de Acessos")
    v["jira_url"] = g("jira/url_chamado")
    v["jira_tipo"] = g("jira/tipo_solicitud", "forms")
    v["jira_cor"] = g("jira/cor_titulo", "Verde-azulado forte")
    v["jira_prefixo"] = g("jira/categoria_prefixo", "JIRA -")
    v["flx_remetente"] = g("fluxo/remetente_filtro", "Microsoft Forms")
    v["flx_subpasta"] = g("fluxo/subpasta_destino", "Finalizados")
    v["flx_so_nao_lidos"] = g("fluxo/so_nao_lidos", "true")
    v["flx_marcar_lido"] = g("fluxo/marcar_como_lido", "true")
    v["flx_encerrar"] = g("fluxo/encerrar_sessoes_no_fim", "true")
    v["flx_dry_run"] = g("fluxo/dry_run", "true")
    v["flx_intervalo"] = g("fluxo/intervalo_min", "5")
    v["flx_invisivel"] = g("fluxo/invisivel", "true")
    v["enc_ativo"] = g("encaminhamento/ativo", "true")
    v["enc_dest"] = g("encaminhamento/destinatarios")
    v["enc_assunto"] = g("encaminhamento/assunto_template", "{ticket} - {assunto}")
    v["enc_prefixo"] = g("encaminhamento/prefixo_corpo", "Chamado aberto: ")
    v["enc_sep"] = g("encaminhamento/separador",
                     "--- Mensagem original (Microsoft Forms) abaixo ---")
    return v


def _b(x):
    return "true" if str(x).strip().lower() in ("true", "1", "sim", "yes") else "false"


# -------------------- gravacao (template com comentarios) --------------------
def _salvar(vals):
    e = escape
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<!--
  Configuracao do projeto CVC - Gestao de Acessos.
  ATENCAO: contem SENHA em texto puro. Manter apenas na maquina local do
  cliente, com acesso restrito. NAO versionar / NAO sincronizar em nuvem.
  (Editado pelo formulario config_form.py)
-->
<config>

  <!-- Conta Microsoft / Outlook -->
  <outlook>
    <email>{e(vals['outlook_email'])}</email>
    <senha>{e(vals['outlook_senha'])}</senha>
    <url>{e(vals['outlook_url'])}</url>
    <caixa>{e(vals['outlook_caixa'])}</caixa>
  </outlook>

  <!-- Jira Service Management -->
  <jira>
    <url_chamado>{e(vals['jira_url'])}</url_chamado>
    <tipo_solicitud>{e(vals['jira_tipo'])}</tipo_solicitud>
    <cor_titulo>{e(vals['jira_cor'])}</cor_titulo>
    <categoria_prefixo>{e(vals['jira_prefixo'])}</categoria_prefixo>
  </jira>

  <!-- Regras do fluxo -->
  <fluxo>
    <remetente_filtro>{e(vals['flx_remetente'])}</remetente_filtro>
    <subpasta_destino>{e(vals['flx_subpasta'])}</subpasta_destino>
    <so_nao_lidos>{_b(vals['flx_so_nao_lidos'])}</so_nao_lidos>
    <marcar_como_lido>{_b(vals['flx_marcar_lido'])}</marcar_como_lido>
    <encerrar_sessoes_no_fim>{_b(vals['flx_encerrar'])}</encerrar_sessoes_no_fim>
    <dry_run>{_b(vals['flx_dry_run'])}</dry_run>
    <intervalo_min>{e(str(vals['flx_intervalo']))}</intervalo_min>
    <invisivel>{_b(vals['flx_invisivel'])}</invisivel>
  </fluxo>

  <!-- Encaminhamento do e-mail pro grupo responsavel (com nº do chamado) -->
  <!-- destinatarios: separados por virgula. Placeholders: {{ticket}} {{assunto}} -->
  <encaminhamento>
    <ativo>{_b(vals['enc_ativo'])}</ativo>
    <destinatarios>{e(vals['enc_dest'])}</destinatarios>
    <assunto_template>{e(vals['enc_assunto'])}</assunto_template>
    <prefixo_corpo>{e(vals['enc_prefixo'])}</prefixo_corpo>
    <!-- separador vazio = usa o divisor automatico do OWA -->
    <separador></separador>
  </encaminhamento>

</config>
"""
    with open(ARQ, "w", encoding="utf-8") as f:
        f.write(xml)


# -------------------- UI --------------------
def main():
    dados = _ler()
    root = tk.Tk()
    root.title("Configuração — CVC Gestão de Acessos")
    root.geometry("560x430")
    root.minsize(520, 380)

    campos = {}       # nome -> tk.StringVar/BooleanVar

    cont = ttk.Frame(root, padding=10)
    cont.pack(fill="both", expand=True)

    nb = ttk.Notebook(cont)
    nb.pack(fill="both", expand=True)

    def aba(titulo):
        fr = ttk.Frame(nb, padding=12)
        nb.add(fr, text=titulo)
        return fr

    def linha(parent, rotulo, chave, senha=False, largura=42):
        fr = ttk.Frame(parent)
        fr.pack(fill="x", pady=4)
        ttk.Label(fr, text=rotulo, width=20).pack(side="left")
        var = tk.StringVar(value=dados.get(chave, ""))
        ent = ttk.Entry(fr, textvariable=var, width=largura,
                        show="*" if senha else "")
        ent.pack(side="left", fill="x", expand=True)
        campos[chave] = var
        if senha:
            mv = tk.BooleanVar(value=False)

            def toggle():
                ent.config(show="" if mv.get() else "*")
            ttk.Checkbutton(fr, text="mostrar", variable=mv,
                            command=toggle).pack(side="left", padx=4)
        return var

    def check(parent, rotulo, chave):
        var = tk.BooleanVar(value=_b(dados.get(chave, "false")) == "true")
        ttk.Checkbutton(parent, text=rotulo, variable=var).pack(anchor="w", pady=2)
        campos[chave] = var
        return var

    # --- Aba Outlook ---
    g = aba("Outlook")
    linha(g, "E-mail", "outlook_email")
    linha(g, "Senha", "outlook_senha", senha=True)
    linha(g, "URL", "outlook_url")
    linha(g, "Caixa compartilhada", "outlook_caixa")

    # --- Aba Jira ---
    g = aba("Jira")
    linha(g, "URL do chamado", "jira_url")
    linha(g, "Tipo de solicitação", "jira_tipo")
    linha(g, "Cor do título", "jira_cor")

    # --- Aba Encaminhamento ---
    g = aba("Encaminhamento")
    check(g, "Encaminhamento ativo", "enc_ativo")
    linha(g, "Destinatários (vírgula)", "enc_dest")
    linha(g, "Assunto (template)", "enc_assunto")
    linha(g, "Prefixo do corpo", "enc_prefixo")
    ttk.Label(g, text="Placeholders do assunto: {ticket}  {assunto}",
              foreground="#666").pack(anchor="w", pady=(6, 0))

    # --- Aba Fluxo ---
    g = aba("Fluxo")
    linha(g, "Remetente (filtro)", "flx_remetente")
    linha(g, "Subpasta destino", "flx_subpasta")
    linha(g, "Intervalo (min)", "flx_intervalo", largura=10)
    check(g, "Só não lidos", "flx_so_nao_lidos")
    check(g, "Marcar como lido", "flx_marcar_lido")
    check(g, "Encerrar sessões no fim", "flx_encerrar")
    check(g, "DRY-RUN (ensaio, não cria/envia nada)", "flx_dry_run")
    check(g, "Invisível (esconde o navegador)", "flx_invisivel")

    # --- Barra de botoes (fixa embaixo, fora das abas) ---
    barra = ttk.Frame(cont)
    barra.pack(fill="x", pady=(10, 0))

    def salvar():
        vals = {}
        for k, var in campos.items():
            vals[k] = var.get()
        # campos nao editaveis no form, mas preservados no config.xml
        vals.setdefault("jira_prefixo", dados.get("jira_prefixo", "JIRA -"))
        vals.setdefault("enc_sep", dados.get("enc_sep", ""))
        try:
            int(str(vals["flx_intervalo"]).strip() or "5")
        except ValueError:
            messagebox.showerror("Erro", "Intervalo (min) precisa ser um número.")
            return
        if not vals["outlook_email"].strip():
            messagebox.showerror("Erro", "E-mail do Outlook é obrigatório.")
            return
        try:
            _salvar(vals)
            messagebox.showinfo("OK", "Configuração salva em config.xml.")
            root.destroy()
        except Exception as ex:  # noqa: BLE001
            messagebox.showerror("Erro ao salvar", str(ex))

    def abrir_xml():
        try:
            os.startfile(ARQ)  # type: ignore[attr-defined]
        except Exception as ex:  # noqa: BLE001
            messagebox.showerror("Erro", str(ex))

    ttk.Button(barra, text="Salvar", command=salvar).pack(side="right", padx=4)
    ttk.Button(barra, text="Cancelar",
               command=root.destroy).pack(side="right", padx=4)
    ttk.Button(barra, text="Abrir config.xml",
               command=abrir_xml).pack(side="left", padx=4)

    root.mainloop()


if __name__ == "__main__":
    main()
