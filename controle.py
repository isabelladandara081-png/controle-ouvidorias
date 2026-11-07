"""Núcleo do Controle de Ouvidorias (sem interface).

Fluxo:  Pendente -> Dilação solicitada -> Controle de Dilatadas -> Acompanhamento
        -> importação da base de Respondidas -> Histórico de respondidas.

Regras de prazo:
  * sem justificativa de dilação: segue o prazo da unidade;
  * com justificativa: "Dilatada" (nunca vencida pelo prazo da unidade);
  * prazo final ultrapassado: "Vencida", mesmo dilatada;
  * só a base de Respondidas conclui uma ouvidoria.
"""
from __future__ import annotations

import json
import os
import re
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import openpyxl
import pandas as pd

AMARELOS = {"FFFF00", "FFC000", "FFD966", "FFE699", "FFEB9C"}
DIAS_PROXIMO = 3


# ----------------------------------------------------------------- utilitários
def so_digitos(v) -> str:
    return re.sub(r"\D", "", "" if v is None else str(v))


def chave_uc(v) -> str:
    m = re.search(r"\d{5,}", "" if v is None else str(v))
    return m.group(0).lstrip("0") if m else ""


def para_iso(v) -> str:
    """Converte datas de vários formatos (datetime, 20261109, 09/11/2026, ISO) para AAAA-MM-DD."""
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if re.fullmatch(r"\d{8}", s):
        return f"{s[:4]}-{s[4:6]}-{s[6:]}"
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", s)
    if m:
        return f"{m[3]}-{m[2]}-{m[1]}"
    m = re.match(r"\d{4}-\d{2}-\d{2}", s)
    return m[0] if m else ""


def dias_ate(iso: str | None, hoje: date | None = None):
    if not iso:
        return None
    return (date.fromisoformat(iso) - (hoje or date.today())).days


def fmt_data(iso: str | None) -> str:
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}" if iso else "—"


def agora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def codigo_da_linha(row, idx=None) -> str:
    """Código normalizado (só dígitos). Procura na coluna indicada e, se falhar,
    em qualquer célula da linha (cobre linhas deslocadas ou com marcadores antes do código)."""
    d = so_digitos(row[idx]) if idx is not None and idx < len(row) else ""
    if len(d) == 12:
        return d
    for x in row:
        if x is not None and re.fullmatch(r"[\d.\s*#\-]+", str(x).strip()) and len(so_digitos(x)) == 12:
            return so_digitos(x)
    return d if len(d) >= 11 else ""


def novo_caso(codigo: str, **kw) -> dict:
    c = {
        "codigo": codigo, "uc": "", "analista": "", "recebimento": "", "assunto": "", "causa": "",
        "unidade": "", "prazo_unidade": "", "prazo_final": "", "forma": "", "instancia": "",
        "justificativa": "", "data_solicitacao": "", "qtd_dilacoes": 0, "aguarda_devolutiva": None,
        "fiscal": "", "oe": "", "equipe": "", "observacoes": "",
        "finalizada": False, "data_finalizacao": "", "atualizado_em": "", "prazo_pessoal": "",
        "sem_base": False,
    }
    c.update(kw)
    return c


# ----------------------------------------------------------------- regras
def esta_dilatada(c: dict) -> bool:
    return (not c.get("finalizada")) and bool(str(c.get("justificativa") or "").strip())


def situacao(c: dict, hoje: date | None = None) -> str:
    if c.get("finalizada"):
        return "respondida"
    f = dias_ate(c.get("prazo_final"), hoje)
    if f is not None and f < 0:
        return "vencida"
    if esta_dilatada(c):
        return "dilatada"
    u = dias_ate(c.get("prazo_unidade"), hoje)
    if u is not None and u < 0:
        return "unidade_vencida"
    return "no_prazo"


ROTULOS = {"respondida": "Respondida", "vencida": "Vencida", "dilatada": "Dilatada",
           "unidade_vencida": "Unidade vencida", "no_prazo": "No prazo"}


def proximo_prazo_final(c: dict, hoje: date | None = None) -> bool:
    f = dias_ate(c.get("prazo_final"), hoje)
    return (not c.get("finalizada")) and f is not None and 0 <= f <= DIAS_PROXIMO


def aguardando_devolutiva(c: dict) -> bool:
    return esta_dilatada(c) and c.get("aguarda_devolutiva") is True


def rotulo_situacao(c: dict, hoje: date | None = None) -> str:
    t = ROTULOS[situacao(c, hoje)]
    if esta_dilatada(c):
        if proximo_prazo_final(c, hoje):
            t += " · Próx. prazo final"
        if aguardando_devolutiva(c):
            t += " · Aguardando devolutiva"
    return t


def _ord(c):
    return (c.get("prazo_unidade") or "9", c.get("prazo_final") or "9")


def _ord_final(c):
    return (c.get("prazo_final") or "9", c.get("prazo_unidade") or "9")


def set_justificativa(c: dict, texto: str, hoje: date | None = None) -> None:
    """Preencher a justificativa = solicitar dilação (o protocolo passa a Dilatada automaticamente)."""
    c["justificativa"] = (texto or "").strip()
    if c["justificativa"]:
        if not c.get("data_solicitacao"):
            c["data_solicitacao"] = (hoje or date.today()).isoformat()
        if not c.get("qtd_dilacoes"):
            c["qtd_dilacoes"] = 1
        if c.get("aguarda_devolutiva") is None:
            c["aguarda_devolutiva"] = True
    c["atualizado_em"] = agora()


# ----------------------------------------------------------------- armazenamento
class Armazem:
    """Guarda os casos (um por código, sem duplicar) em dados/ouvidorias.json."""

    def __init__(self, pasta: str | Path | None = None):
        self.pasta = Path(pasta or os.environ.get("OUVIDORIAS_DADOS", "dados"))
        self.arq = self.pasta / "ouvidorias.json"
        self.bak = self.pasta / "antes_da_ultima_importacao.json"
        self.casos: dict[str, dict] = {}
        self.meta: dict = {}
        if self.arq.exists():
            d = json.loads(self.arq.read_text(encoding="utf-8"))
            self.casos, self.meta = d.get("casos", {}), d.get("meta", {})

    def _dump(self) -> str:
        return json.dumps({"versao": 1, "casos": self.casos, "meta": self.meta}, ensure_ascii=False, indent=1)

    def salvar(self) -> None:
        self.pasta.mkdir(parents=True, exist_ok=True)
        tmp = self.arq.with_suffix(".tmp")
        tmp.write_text(self._dump(), encoding="utf-8")
        os.replace(tmp, self.arq)

    def ponto_de_restauracao(self) -> None:
        self.pasta.mkdir(parents=True, exist_ok=True)
        self.bak.write_text(self._dump(), encoding="utf-8")

    def desfazer(self) -> bool:
        if not self.bak.exists():
            return False
        d = json.loads(self.bak.read_text(encoding="utf-8"))
        self.casos, self.meta = d["casos"], d.get("meta", {})
        self.salvar()
        return True

    def exportar_json(self) -> str:
        return self._dump()

    def restaurar_json(self, texto: str) -> int:
        d = json.loads(texto)
        if not isinstance(d.get("casos"), dict):
            raise ValueError("Arquivo de backup inválido.")
        self.ponto_de_restauracao()
        self.casos, self.meta = d["casos"], d.get("meta", {})
        self.salvar()
        return len(self.casos)

    def limpar(self) -> None:
        self.ponto_de_restauracao()
        self.casos, self.meta = {}, {}
        self.salvar()

    # listas por área
    def pendentes(self):
        return sorted((c for c in self.casos.values() if not c["finalizada"] and not esta_dilatada(c)), key=_ord)

    def dilatadas(self):
        return sorted((c for c in self.casos.values() if esta_dilatada(c)), key=_ord_final)

    def respondidas(self):
        return sorted((c for c in self.casos.values() if c["finalizada"]),
                      key=lambda c: c.get("data_finalizacao") or "", reverse=True)


CAMPOS_EDITAVEIS = {"uc", "analista", "assunto", "prazo_unidade", "prazo_final", "justificativa", "fiscal", "oe",
                    "equipe", "qtd_dilacoes", "aguarda_devolutiva", "observacoes", "data_solicitacao"}


def atualizar(arm: Armazem, codigo: str, **campos) -> None:
    """Atualiza campos de um caso e salva. A movimentação entre abas é consequência das regras."""
    c = arm.casos[codigo]
    for k, v in campos.items():
        if k not in CAMPOS_EDITAVEIS:
            raise KeyError(k)
        if k == "justificativa":
            set_justificativa(c, v)
        else:
            if k == "qtd_dilacoes":
                v = int(v) if str(v).strip() else 0
            c[k] = v
            c["atualizado_em"] = agora()
    arm.salvar()


# ----------------------------------------------------------------- painel
def prioridades(arm: Armazem, hoje: date | None = None) -> dict:
    P = [c for c in arm.casos.values() if not c["finalizada"]]
    f = lambda c: dias_ate(c.get("prazo_final"), hoje)  # noqa: E731
    u = lambda c: dias_ate(c.get("prazo_unidade"), hoje)  # noqa: E731
    normais = [c for c in P if not esta_dilatada(c) and not (f(c) is not None and f(c) < 0) and u(c) is not None]
    return {
        "proximas_prazo_final": sorted((c for c in P if proximo_prazo_final(c, hoje)), key=_ord_final),
        "hoje": sorted((c for c in normais if u(c) <= 0), key=_ord),
        "amanha": sorted((c for c in normais if u(c) == 1), key=_ord),
        "proximo": sorted((c for c in normais if 2 <= u(c) <= DIAS_PROXIMO), key=_ord),
        "final_vencido": sorted((c for c in P if f(c) is not None and f(c) < 0), key=_ord_final),
    }


def _rel(d):
    return "vencido" if d < 0 else "vence hoje" if d == 0 else "amanhã" if d == 1 else f"em {d} dias"


def linha_tabela(c: dict, hoje: date | None = None) -> dict:
    u = dias_ate(c.get("prazo_unidade"), hoje)
    f = dias_ate(c.get("prazo_final"), hoje)
    pu = fmt_data(c.get("prazo_unidade"))
    if u is not None and not c["finalizada"] and not esta_dilatada(c):
        pu += f" ({_rel(u)})"
    pf = fmt_data(c.get("prazo_final"))
    if f is not None and not c["finalizada"] and f <= DIAS_PROXIMO:
        pf += f" ({_rel(f)})"
    return {"Código": c["codigo"], "UC": c.get("uc", ""), "Analista": c.get("analista", ""),
            "Assunto": c.get("assunto") or c.get("causa", ""), "Prazo da unidade": pu, "Prazo final": pf,
            "Situação": rotulo_situacao(c, hoje)}


# ----------------------------------------------------------------- importações
def _mapear(vals, padroes):
    m = {}
    for i, h in enumerate(vals):
        for k, p in padroes.items():
            if k not in m and re.search(p, h, re.I):
                m[k] = i
                break
    return m


def _cabecalho(ws, teste):
    for n, row in enumerate(ws.iter_rows(values_only=True), start=1):
        vals = ["" if v is None else str(v).strip() for v in row]
        if teste(vals):
            return n, vals
    return None


def _amarela(ws, n) -> bool:
    for cel in ws[n]:
        f = cel.fill
        if f is not None and f.fill_type == "solid" and f.fgColor.type == "rgb":
            rgb = f.fgColor.rgb
            if isinstance(rgb, str) and rgb[-6:].upper() in AMARELOS:
                return True
    return False


def _t(row, i):
    return "" if i is None or i >= len(row) or row[i] is None else str(row[i]).strip()


def _abrir(arquivo):
    return openpyxl.load_workbook(arquivo, data_only=True)


def importar_aguardando(arm: Armazem, arquivo) -> str:
    """Base diária 'Aguardando resposta': entram códigos novos, os existentes são atualizados
    preservando o que foi preenchido à mão. Respondidas não são reabertas."""
    arm.ponto_de_restauracao()
    pad = {"codigo": r"^c[oó]digo", "recebimento": r"^recebimento", "unidade": r"^unidade respons",
           "prazo_unidade": r"^prazo da unidade", "prazo_final": r"^prazo final", "forma": r"^forma",
           "instancia": r"^inst", "assunto": r"^subassunto", "analista": r"^analista"}
    novas = atual = ignoradas = 0
    achou = False
    for ws in _abrir(arquivo).worksheets:
        cab = _cabecalho(ws, lambda v: any(re.match(r"c[oó]digo", x, re.I) for x in v))
        if not cab:
            continue
        achou = True
        n0, vals = cab
        m = _mapear(vals, pad)
        for row in ws.iter_rows(min_row=n0 + 1, values_only=True):
            cod = codigo_da_linha(row, m.get("codigo"))
            if not cod:
                continue
            dados = {"recebimento": para_iso(row[m["recebimento"]]) if "recebimento" in m else "",
                     "unidade": _t(row, m.get("unidade")),
                     "prazo_unidade": para_iso(row[m["prazo_unidade"]]) if "prazo_unidade" in m else "",
                     "prazo_final": para_iso(row[m["prazo_final"]]) if "prazo_final" in m else "",
                     "forma": _t(row, m.get("forma")), "instancia": _t(row, m.get("instancia")),
                     "assunto": _t(row, m.get("assunto")), "analista": _t(row, m.get("analista"))}
            c = arm.casos.get(cod)
            if c and c["finalizada"] and not c.get("sem_base"):
                ignoradas += 1
                continue
            if c:
                if c["finalizada"]:  # criada só pelo acompanhamento e agora apareceu na base: reabre
                    c["finalizada"], c["sem_base"] = False, False
                for k, v in dados.items():
                    if v != "" or not c.get(k):
                        c[k] = v
                atual += 1
            else:
                arm.casos[cod] = novo_caso(cod, **dados)
                novas += 1
    if not achou:
        raise ValueError("Não encontrei a coluna 'Código'. Confira se é a base 'Aguardando resposta'.")
    arm.meta["ultima_base_diaria"] = agora()
    arm.salvar()
    return f"Aguardando resposta: {novas} nova(s), {atual} atualizada(s), {ignoradas} ignorada(s) por já estarem respondidas."


def importar_respondidas(arm: Armazem, arquivo) -> str:
    """Base de Ouvidorias Respondidas: conclui automaticamente as pendentes/dilatadas encontradas
    (preservando UC, observações, justificativa, OE e fiscal). Códigos novos vão só ao histórico."""
    arm.ponto_de_restauracao()
    vistos, concl, ja, novos = set(), 0, 0, 0
    wb = _abrir(arquivo)
    abas = []  # só abas com cabeçalho de código; se nenhuma tiver, lê apenas a primeira aba
    for ws in wb.worksheets:
        cab = _cabecalho(ws, lambda v: any(re.search(r"c[oó]digo|protocolo|ouvidoria", x, re.I) for x in v))
        if cab:
            m = _mapear(cab[1], {"cod": r"c[oó]digo|protocolo|ouvidoria$", "dt": r"resposta|conclus"})
            if "cod" in m:
                abas.append((ws, cab[0], m))
    if not abas:
        abas = [(wb.worksheets[0], 0, {})]
    for ws, n0, m in abas:
        for row in ws.iter_rows(min_row=n0 + 1, values_only=True):
            cod = codigo_da_linha(row, m.get("cod"))
            if not cod or cod in vistos:
                continue
            vistos.add(cod)
            dt = para_iso(row[m["dt"]]) if "dt" in m and m["dt"] < len(row) else ""
            fin = f"{dt}T12:00:00" if dt else agora()
            c = arm.casos.get(cod)
            if c is None:
                arm.casos[cod] = novo_caso(cod, finalizada=True, data_finalizacao=fin)
                novos += 1
            elif c["finalizada"]:
                ja += 1
            else:
                c.update(finalizada=True, data_finalizacao=fin, atualizado_em=agora())
                concl += 1
    if not vistos:
        raise ValueError("Não encontrei nenhum código de 12 dígitos nesta planilha.")
    arm.meta["ultima_respondidas"] = agora()
    arm.salvar()
    return (f"Respondidas: {len(vistos)} código(s) lidos. {concl} pendente(s)/dilatada(s) concluída(s) "
            f"automaticamente, {ja} já estavam no histórico, {novos} nova(s) entraram só no histórico.")


def importar_acompanhamento(arm: Armazem, arquivo) -> str:
    """Planilha de acompanhamento: vincula UC, OE, equipe, fiscal e observações pelo código.
    Linhas amarelas ou com 'dilação' entram como dilatadas. Nunca sobrescreve o que já foi preenchido."""
    arm.ponto_de_restauracao()
    pad = {"data": r"^data", "ouv": r"^ouvidoria", "oe": r"^oe", "uc": r"^uc", "eq": r"^equipe", "obs": r"^obs"}
    novos = vinc = dilat = 0
    achou = False
    for ws in _abrir(arquivo).worksheets:
        cab = _cabecalho(ws, lambda v: any(re.match(r"ouvidoria", x, re.I) for x in v)
                         and any(re.match(r"uc", x, re.I) for x in v))
        if not cab:
            continue
        achou = True
        n0, vals = cab
        m = _mapear(vals, pad)
        for n, row in enumerate(ws.iter_rows(min_row=n0 + 1, values_only=True), start=n0 + 1):
            if not any(x is not None and str(x).strip() for x in row):
                continue
            uc, oe, eq, ob = (_t(row, m.get(k)) for k in ("uc", "oe", "eq", "obs"))
            texto_data = _t(row, m.get("data"))
            dt = para_iso(row[m["data"]]) if "data" in m and m["data"] < len(row) else ""
            cod = codigo_da_linha(row, m.get("ouv"))
            if not cod:
                k = chave_uc(uc)
                if not k:
                    continue
                cod = f"SEM CÓDIGO UC {k} {(oe or eq or texto_data).upper()}"[:60]
            dila = _amarela(ws, n) or bool(re.search("dila", " ".join(str(x) for x in row if x is not None), re.I))
            c = arm.casos.get(cod)
            if c is None:
                c = novo_caso(cod, manual=True)
                if not dila:  # fora da base atual e sem dilação: fica só no histórico
                    c["finalizada"], c["sem_base"] = True, True
                arm.casos[cod] = c
                novos += 1
            else:
                vinc += 1

            def preencher(k, v):
                if v and not c.get(k):
                    c[k] = v

            preencher("uc", uc)
            preencher("recebimento", dt)
            obs = " | ".join(x for x in (ob or ("" if dt else texto_data),
                                         oe if oe and not oe.isdigit() and "fiscal" not in oe.lower() else "") if x)
            preencher("observacoes", obs)
            if oe.isdigit():
                preencher("oe", oe)
            if re.search("fiscal", f"{oe} {eq}", re.I):
                preencher("fiscal", oe or eq)
            if eq and not re.search("fiscal|dila", eq, re.I):
                preencher("equipe", eq)
            if dila and c.get("sem_base"):
                c["finalizada"], c["sem_base"] = False, False
            if dila and not esta_dilatada(c) and not c["finalizada"]:
                set_justificativa(c, ob or oe or eq or "Dilatada (planilha de acompanhamento)")
                dilat += 1
    if not achou:
        raise ValueError("Não encontrei as colunas OUVIDORIA e UC. Confira se é a planilha de acompanhamento.")
    arm.meta["ultimo_acompanhamento"] = agora()
    arm.salvar()
    return f"Acompanhamento: {novos} registro(s) novo(s), {vinc} vinculado(s) a existentes, {dilat} entraram como dilatadas."


def importar_backup_html(arm: Armazem, texto: str) -> int:
    """Traz os dados do aplicativo em HTML (backup .json) sem apagar nada que já exista aqui."""
    d = json.loads(texto)
    arm.ponto_de_restauracao()
    n = 0
    for o in d.get("cases", []):
        cod = so_digitos(o.get("codigo")) or str(o.get("codigo") or "")
        if not cod:
            continue
        j = o.get("justificativa") or (o.get("motivo") if o.get("dilatada") else "") or ""
        dados = {"uc": o.get("uc", ""), "analista": o.get("analista", ""), "recebimento": o.get("recebimento", ""),
                 "assunto": o.get("assunto", ""), "causa": o.get("causa", ""), "unidade": o.get("unidade", ""),
                 "prazo_unidade": o.get("prazoUnidade", ""), "prazo_final": o.get("prazoFinal", ""),
                 "forma": o.get("forma", ""), "instancia": o.get("instancia", ""), "justificativa": j,
                 "data_solicitacao": o.get("dilData", ""), "qtd_dilacoes": int(o.get("dilSolic") or (1 if j else 0)),
                 "aguarda_devolutiva": (o.get("aguardaDev") == "Sim") if o.get("aguardaDev") else None,
                 "fiscal": o.get("fiscal", ""), "oe": o.get("oeNumero", ""), "equipe": o.get("equipe", ""),
                 "observacoes": o.get("obs", ""), "finalizada": bool(o.get("finalizada")),
                 "data_finalizacao": o.get("dataFinal", ""), "atualizado_em": o.get("upd", ""),
                 "prazo_pessoal": o.get("prazoPessoal", "")}
        c = arm.casos.get(cod)
        if c is None:
            arm.casos[cod] = novo_caso(cod, **dados)
            n += 1
        else:
            for k, v in dados.items():
                if v not in ("", None) and not c.get(k):
                    c[k] = v
    arm.salvar()
    return n


# ----------------------------------------------------------------- exportação
def para_dataframe(casos, hoje: date | None = None) -> pd.DataFrame:
    linhas = []
    for c in casos:
        linhas.append({
            "Código": c["codigo"], "UC": c.get("uc", ""), "Analista": c.get("analista", ""),
            "Recebimento": fmt_data(c.get("recebimento")), "Assunto": c.get("assunto") or c.get("causa", ""),
            "Prazo da unidade": fmt_data(c.get("prazo_unidade")), "Prazo final": fmt_data(c.get("prazo_final")),
            "Situação": rotulo_situacao(c, hoje), "Justificativa da dilação": c.get("justificativa", ""),
            "Data da solicitação": fmt_data(c.get("data_solicitacao")), "Dilatações solicitadas": c.get("qtd_dilacoes") or "",
            "Aguardando devolutiva": "Sim" if c.get("aguarda_devolutiva") else "Não",
            "Fiscal acionado": c.get("fiscal", ""), "OE aberta": c.get("oe", ""), "Equipe acionada": c.get("equipe", ""),
            "Observações e retornos": c.get("observacoes", ""),
            "Respondida em": fmt_data((c.get("data_finalizacao") or "")[:10]),
            "Última atualização": c.get("atualizado_em", "")})
    return pd.DataFrame(linhas)


def exportar_excel(arm: Armazem, so_dilatadas: bool = False) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        if so_dilatadas:
            para_dataframe(arm.dilatadas()).to_excel(w, sheet_name="Dilatadas", index=False)
        else:
            para_dataframe(arm.pendentes()).to_excel(w, sheet_name="Pendentes", index=False)
            para_dataframe(arm.dilatadas()).to_excel(w, sheet_name="Dilatadas", index=False)
            para_dataframe(arm.respondidas()).to_excel(w, sheet_name="Respondidas", index=False)
    return buf.getvalue()
