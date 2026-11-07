
"""Controle de Ouvidorias — interface Streamlit.  Execute com:  streamlit run app.py"""
from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

import controle as oc

st.set_page_config(page_title="OUVIDORIAS", page_icon="📞", layout="wide")
arm = oc.Armazem()
st.session_state.setdefault("v", 0)

COLUNAS_TABELA = ["Código", "UC", "Analista", "Assunto", "Prazo da unidade", "Prazo final", "Situação"]
CAMPO = {"UC": "uc", "Justificativa da dilação": "justificativa", "Fiscal acionado": "fiscal", "OE aberta": "oe",
         "Equipe acionada": "equipe", "Dilatações solicitadas": "qtd_dilacoes",
         "Aguardando devolutiva": "aguarda_devolutiva", "Observações e retornos": "observacoes"}


def _n(v):
    """Normaliza valores vindos do editor (NaN/None -> '')."""
    if v is None or (isinstance(v, float) and v != v):
        return ""
    return v.item() if hasattr(v, "item") else v


def tabela(casos, **kw):
    if not casos:
        st.caption("Nenhuma ouvidoria.")
        return
    st.dataframe(pd.DataFrame([oc.linha_tabela(c) for c in casos])[COLUNAS_TABELA],
                 hide_index=True, use_container_width=True, **kw)


def editar(df: pd.DataFrame, editaveis: list[str], key: str, config=None):
    """data_editor que grava cada alteração. A chave muda a cada gravação para evitar
    que edições antigas sejam aplicadas em linhas que mudaram de posição."""
    ed = st.data_editor(df, hide_index=True, use_container_width=True, key=f"{key}_{st.session_state.v}",
                        disabled=[c for c in df.columns if c not in editaveis], column_config=config or {})
    mudou = False
    for i in range(len(df)):
        cod = df.iloc[i]["Código"]
        for col in editaveis:
            antes, depois = _n(df.iloc[i][col]), _n(ed.iloc[i][col])
            if antes != depois:
                oc.atualizar(arm, cod, **{CAMPO[col]: depois})
                mudou = True
    if mudou:
        st.session_state.v += 1
        st.rerun()


def filtrar(casos, busca: str):
    b = busca.strip().lower()
    return [c for c in casos if not b or b in f"{c['codigo']} {c.get('uc', '')} {c.get('analista', '')} "
                                             f"{c.get('assunto', '')} {c.get('causa', '')}".lower()]


st.title("📞 Controle de Ouvidorias")
t_painel, t_pend, t_dil, t_hist, t_imp, t_bkp = st.tabs(
    ["📊 Painel de prioridades", "📋 Ouvidorias", "⏳ Controle de Dilatadas", "✅ Histórico de respondidas",
     "📥 Importar bases", "💾 Backup e exportação"])

# ---------------------------------------------------------------- Painel
with t_painel:
    pr = oc.prioridades(arm)
    with st.container(border=True):
        st.markdown(f"### ⚠️ Próximas do prazo final (até {oc.DIAS_PROXIMO} dias) — {len(pr['proximas_prazo_final'])}")
        st.caption("Inclui as dilatadas. O prazo final é o que determina o vencimento real.")
        tabela(pr["proximas_prazo_final"])
    for titulo, chave, icone in [("1. Vencendo hoje (prazo da unidade)", "hoje", "🔴"),
                                 ("2. Vencendo amanhã", "amanha", "🟠"),
                                 ("3. Prazo próximo (2 a 3 dias)", "proximo", "🔵"),
                                 ("5. Prazo final vencido", "final_vencido", "🔴")]:
        st.markdown(f"#### {icone} {titulo} — {len(pr[chave])}")
        tabela(pr[chave])
    st.caption("O item 4 (dilatadas próximas do prazo final) está na caixa em destaque acima.")

# ---------------------------------------------------------------- Ouvidorias pendentes
with t_pend:
    st.subheader("Ouvidorias pendentes de resposta")
    st.caption("Preencha a **Justificativa da dilação** para transferir a ouvidoria automaticamente "
               "para o Controle de Dilatadas. Dê Enter para gravar.")
    busca = st.text_input("Buscar por código, UC, analista ou assunto", key="b_pend")
    casos = filtrar(arm.pendentes(), busca)
    st.caption(f"{len(casos)} pendente(s)")
    if casos:
        df = pd.DataFrame([{**oc.linha_tabela(c), "Justificativa da dilação": ""} for c in casos])
        editar(df, ["UC", "Justificativa da dilação"], "ed_pend")

# ---------------------------------------------------------------- Controle de Dilatadas
with t_dil:
    st.subheader("Controle de Dilatadas")
    c1, c2, c3 = st.columns([2, 1, 1])
    busca = c1.text_input("Buscar por código ou UC", key="b_dil")
    usar_data = c2.checkbox("Filtrar por prazo final")
    ate = c2.date_input("Prazo final até", value=date.today(), disabled=not usar_data)
    so_ag = c3.checkbox("Só aguardando devolutiva")
    todas = arm.dilatadas()
    b = busca.strip().lower()
    lista = [c for c in todas
             if (not b or b in f"{c['codigo']} {c.get('uc', '')}".lower())
             and (not usar_data or (c.get("prazo_final") or "9") <= ate.isoformat())
             and (not so_ag or oc.aguardando_devolutiva(c))]

    aguard = [c for c in todas if oc.aguardando_devolutiva(c)]
    with st.container(border=True):
        st.markdown(f"### 📨 Aguardando devolutiva — {len(aguard)}")
        st.caption("Cobrar a devolução da ouvidoria com o analista responsável.")
        for c in sorted(aguard, key=lambda c: (c.get("analista") or "~", c.get("prazo_final") or "9")):
            st.markdown(f"**:blue[{(c.get('analista') or 'SEM ANALISTA').upper()}]** — {c['codigo']} · "
                        f"{c.get('qtd_dilacoes') or 1}ª dilação · solicitada em {oc.fmt_data(c.get('data_solicitacao'))}"
                        f" · prazo final {oc.fmt_data(c.get('prazo_final'))}")
        if not aguard:
            st.caption("Nenhuma.")

    st.caption(f"{len(lista)} dilatada(s) · da mais urgente para a menos urgente (prazo final). "
               "🔴 vencida · 🟡 próxima do prazo final. Edite direto na tabela; apagar a justificativa devolve a ouvidoria às pendentes.")
    if lista:
        linhas = []
        for c in lista:
            s = oc.situacao(c)
            icone = "🔴 " if s == "vencida" else "🟡 " if oc.proximo_prazo_final(c) else ""
            linhas.append({"Analista": c.get("analista", ""), "Código": c["codigo"], "UC": c.get("uc", ""),
                           "Situação": icone + oc.rotulo_situacao(c),
                           "Justificativa da dilação": c.get("justificativa", ""),
                           "Data da solicitação": oc.fmt_data(c.get("data_solicitacao")),
                           "Prazo final": oc.fmt_data(c.get("prazo_final")),
                           "Dilatações solicitadas": int(c.get("qtd_dilacoes") or 1),
                           "Aguardando devolutiva": bool(c.get("aguarda_devolutiva")),
                           "Fiscal acionado": c.get("fiscal", ""), "OE aberta": c.get("oe", ""),
                           "Equipe acionada": c.get("equipe", ""),
                           "Observações e retornos": c.get("observacoes", "")})
        editar(pd.DataFrame(linhas), ["UC", "Justificativa da dilação", "Dilatações solicitadas",
               "Aguardando devolutiva", "Fiscal acionado", "OE aberta", "Equipe acionada", "Observações e retornos"],
               "ed_dil", {"Dilatações solicitadas": st.column_config.NumberColumn(min_value=1, step=1)})

# ---------------------------------------------------------------- Histórico
with t_hist:
    st.subheader("Histórico de respondidas")
    busca = st.text_input("Buscar por código, UC, analista ou assunto", key="b_hist")
    r = filtrar(arm.respondidas(), busca)
    st.caption(f"{len(r)} respondida(s)")
    if r:
        df = oc.para_dataframe(r)[["Código", "UC", "Analista", "Assunto", "Respondida em", "Justificativa da dilação",
                                   "Fiscal acionado", "OE aberta", "Equipe acionada", "Observações e retornos"]]
        st.dataframe(df, hide_index=True, use_container_width=True)

# ---------------------------------------------------------------- Importação
def bloco_importacao(titulo, texto, funcao, chave):
    with st.container(border=True):
        st.markdown(f"**{titulo}**")
        st.caption(texto)
        f = st.file_uploader("Planilha (.xlsx)", type=["xlsx"], key=f"up_{chave}")
        if f and st.button("Importar", key=f"bt_{chave}"):
            try:
                st.success(funcao(arm, f))
                st.session_state.v += 1
            except Exception as e:  # noqa: BLE001
                st.error(f"Erro: {e}")


with t_imp:
    st.subheader("Importar bases")
    bloco_importacao("1. Aguardando resposta (diária)",
                     "Códigos novos entram como pendentes; os existentes são atualizados sem perder o que você preencheu.",
                     oc.importar_aguardando, "ag")
    bloco_importacao("2. Ouvidorias respondidas",
                     "As pendentes e dilatadas encontradas são concluídas e vão para o Histórico, com tudo preservado.",
                     oc.importar_respondidas, "rs")
    bloco_importacao("3. Acompanhamento",
                     "Vincula UC, OE, equipe, fiscal e observações pelo código. Linhas amarelas ou com “dilação” entram como dilatadas.",
                     oc.importar_acompanhamento, "ac")

# ---------------------------------------------------------------- Backup
with t_bkp:
    st.subheader("Backup e exportação")
    st.caption(f"{len(arm.casos)} registro(s) em `{arm.arq}`. Um ponto de restauração é criado antes de cada importação.")
    a, b_, c_ = st.columns(3)
    a.download_button("⬇️ Backup completo (.json)", arm.exportar_json(), f"backup-ouvidorias-{date.today()}.json")
    b_.download_button("⬇️ Exportar tudo (Excel)", oc.exportar_excel(arm), f"ouvidorias-{date.today()}.xlsx")
    c_.download_button("⬇️ Exportar dilatadas (Excel)", oc.exportar_excel(arm, True), f"dilatadas-{date.today()}.xlsx")
    if st.button("↩️ Desfazer última importação"):
        st.success("Desfeito.") if arm.desfazer() else st.warning("Não há ponto de restauração.")
        st.session_state.v += 1
    r = st.file_uploader("Restaurar backup (.json)", type=["json"], key="rest")
    if r and st.button("Restaurar este backup"):
        try:
            st.success(f"{arm.restaurar_json(r.getvalue().decode('utf-8'))} registro(s) restaurados.")
            st.session_state.v += 1
        except Exception as e:  # noqa: BLE001
            st.error(str(e))
    h = st.file_uploader("Trazer dados do aplicativo em HTML (backup .json)", type=["json"], key="html")
    if h and st.button("Importar backup do app HTML"):
        try:
            st.success(f"{oc.importar_backup_html(arm, h.getvalue().decode('utf-8'))} registro(s) novo(s).")
            st.session_state.v += 1
        except Exception as e:  # noqa: BLE001
            st.error(str(e))
    with st.expander("Zona de risco"):
        if st.checkbox("Confirmo que quero apagar todos os dados") and st.button("Apagar tudo"):
            arm.limpar()
            st.session_state.v += 1
            st.rerun()
