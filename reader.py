"""
ContábilFlow Python — Centralizador Seguro de Folhas de Pagamento (LGPD)
Evolução do app_local.py com:
  1. Pipeline em etapas reais (st.status)
  2. Extração 100% Local (pdfplumber + Regex) OU Agente Local (PydanticAI + Ollama)
  3. Auditoria aritmética automática e mascaramento de CPF (LGPD)
  4. Chat contábil local sobre o lote extraído (st.chat_message / st.chat_input)
  5. Exportação para CSV (padrão Excel PT-BR) e Relatório TXT

Instalação recomendada (100% on-premise):
    pip install streamlit pandas pdfplumber pydantic
    # Opcional (OCR para PDFs escaneados): pip install pytesseract pillow
    # Opcional (IA 100% local sem internet): pip install pydantic-ai && ollama pull qwen2.5:14b
"""
from __future__ import annotations  # permite "str | re.Pattern" etc. em Python 3.8/3.9

import platform
import re
from typing import List, Optional

import streamlit as st

# Deve ser o 1º comando do Streamlit: garante que algo apareça na tela mesmo se
# alguma dependência abaixo falhar.
st.set_page_config(page_title="ContábilFlow Local", layout="wide")

try:
    import pandas as pd
    import pdfplumber
    from pydantic import BaseModel, Field, model_validator
except ImportError as e:
    st.error(
        f"Dependência ausente ou desatualizada: {e}\n\n"
        "Rode no terminal:  pip install -U streamlit pandas pdfplumber pydantic"
    )
    st.stop()

try:
    import pytesseract
    from PIL import Image
    OCR_DISPONIVEL = True
except ImportError:
    OCR_DISPONIVEL = False

TOLERANCIA = 0.01
MONEY_RE = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}")

LABELS_NUM = {
    "salario_bruto": r"sal[aá]rio\s+bruto|total\s+(?:de\s+)?(?:proventos|vencimentos)|total\s+bruto|sal[aá]rio\s+base(?:\s+mensal)?|sal\.?\s*base(?:\s+mensal)?|sal[aá]rio\s+mensal|sal[aá]rio\s+contratual|vencimento\s+base",
    "inss": r"\binss\b",
    "irrf": r"\birrf\b|imposto\s+de\s+renda",
    "fgts": r"\bfgts\b",
    "total_descontos": r"total\s+(?:de\s+)?descontos",
    "salario_liquido": r"l[ií]quido(?:\s+a\s+receber)?",
}

SALARIO_BRUTO_PRIORITARIO = re.compile(
    r"sal[aá]rio\s+bruto|total\s+(?:de\s+)?(?:proventos|vencimentos)|total\s+bruto",
    re.I,
)

TRES_TOTAIS_RE = re.compile(
    r"total\s+(?:de\s+)?(?:proventos|vencimentos)(?:(?!\d{1,3}(?:\.\d{3})*,\d{2})[\s\S]){0,120}?total\s+(?:de\s+)?descontos(?:(?!\d{1,3}(?:\.\d{3})*,\d{2})[\s\S]){0,120}?(?:valor\s+l[ií]quido|l[ií]quido(?:\s+a\s+receber)?)[\s\S]*?(\d{1,3}(?:\.\d{3})*,\d{2})[\s\S]*?(\d{1,3}(?:\.\d{3})*,\d{2})[\s\S]*?(\d{1,3}(?:\.\d{3})*,\d{2})",
    re.I,
)

RUBRICA_PROVENTO_RE = re.compile(
    r"sal[aá]rio\s+base(?:\s+mensal)?|sal\.?\s*base(?:\s+mensal)?|sal[aá]rio\s+mensal|sal[aá]rio\s+normal|horas?\s+extras?|\bdsr\b|repouso\s+remunerado|adicional|comiss[aãõ]e?s?|gratifica[çc][aã]o|pr[eê]mio|abono|f[eé]rias",
    re.I,
)

LINHA_NAO_PROVENTO_RE = re.compile(
    r"\b(?:inss|irrf|fgts|desconto|descontos|l[ií]quido|base\s+c[aá]lc|sal\.?\s*contr|total\s+de)\b",
    re.I,
)

FGTS_DO_MES_RE = re.compile(r"fgts\s+do\s+m[eê]s[\s\S]*?(\d{1,3}(?:\.\d{3})*,\d{2})", re.I)

NOME_ANTES_CPF_RE = re.compile(
    r"([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ]{2,}(?:[ \t]+[A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ.]{1,}){1,7})\s*(?:CPF\s*[:\-]?\s*)?\d{3}\.\d{3}\.\d{3}-\d{2}"
)

LABELS_TXT = {
    "nome": r"(?m)nome(?:\s+do\s+funcion[aá]rio)?\s*[:\-]\s*(.+?)(?=\s+(?:ref\.?|cpf|cargo|compet[eê]ncia)\b|$)",
    "cpf": r"(\d{3}\.\d{3}\.\d{3}-\d{2})",
    "cargo": r"(?m)cargo\s*[:\-]\s*(.+?)(?=\s+(?:ref\.?|cpf|compet[eê]ncia)\b|$)",
    "competencia": r"(\d{2}/\d{4})",
}


class HoleriteAuditado(BaseModel):
    arquivo: str
    pagina: int
    nome: str = ""
    cpf: str = ""
    cargo: str = ""
    competencia: str = ""
    salario_bruto: float = Field(default=0.0, ge=0.0)
    inss: float = Field(default=0.0, ge=0.0)
    irrf: float = Field(default=0.0, ge=0.0)
    fgts: float = Field(default=0.0, ge=0.0)
    total_descontos: float = Field(default=0.0, ge=0.0)
    salario_liquido: float = Field(default=0.0, ge=0.0)
    divergencia: bool = False
    diferenca: float = 0.0

    @model_validator(mode="after")
    def conferir_totais(self):
        esperado = round(self.salario_bruto - self.total_descontos, 2)
        self.diferenca = round(self.salario_liquido - esperado, 2)
        # Bruto zerado = documento não lido: também deve ser sinalizado para revisão.
        self.divergencia = abs(self.diferenca) > TOLERANCIA or self.salario_bruto == 0
        return self


def br_para_float(valor: str) -> float:
    return float(valor.replace(".", "").replace(",", "."))


def mascarar_cpf(cpf: str) -> str:
    return re.sub(r"\d{3}\.\d{3}\.\d{3}-(\d{2})", r"•••.•••.•••-\1", cpf or "")


def extrair_texto_pagina(pagina) -> str:
    txt = pagina.extract_text() or ""
    if len(txt.strip()) < 30 and OCR_DISPONIVEL:
        img = pagina.to_image(resolution=300).original
        txt = pytesseract.image_to_string(img, lang="por")
    return txt


def _extrair_valor_linha_ou_proxima(linhas: list[str], padrao: str | re.Pattern) -> Optional[float]:
    rx = re.compile(padrao, re.I) if isinstance(padrao, str) else padrao
    for i, ln in enumerate(linhas):
        if rx.search(ln):
            valores = MONEY_RE.findall(ln)
            if valores:
                return br_para_float(valores[-1])
            if i + 1 < len(linhas):
                prox_valores = MONEY_RE.findall(linhas[i + 1])
                if prox_valores:
                    return br_para_float(prox_valores[-1])
    return None


def interpretar_local(arquivo: str, pagina: int, texto: str) -> HoleriteAuditado:
    dados: dict = {"arquivo": arquivo, "pagina": pagina}
    linhas = texto.splitlines()
    for campo in LABELS_NUM:
        dados[campo] = 0.0

    m_totais = TRES_TOTAIS_RE.search(texto)
    if m_totais:
        dados["salario_bruto"] = br_para_float(m_totais.group(1))
        dados["total_descontos"] = br_para_float(m_totais.group(2))
        dados["salario_liquido"] = br_para_float(m_totais.group(3))

    for campo, padrao in LABELS_NUM.items():
        if campo == "fgts":
            m_fgts = FGTS_DO_MES_RE.search(texto)
            if m_fgts:
                dados["fgts"] = br_para_float(m_fgts.group(1))
                continue
        if dados[campo] > 0:
            continue
        if campo == "salario_bruto":
            prioritario = _extrair_valor_linha_ou_proxima(linhas, SALARIO_BRUTO_PRIORITARIO)
            if prioritario is not None:
                dados[campo] = prioritario
                continue
        encontrado = _extrair_valor_linha_ou_proxima(linhas, padrao)
        if encontrado is not None:
            dados[campo] = encontrado

    if dados["salario_bruto"] == 0:
        dados["salario_bruto"] = round(
            sum(
                br_para_float(MONEY_RE.findall(ln)[-1])
                for ln in linhas
                if RUBRICA_PROVENTO_RE.search(ln)
                and not LINHA_NAO_PROVENTO_RE.search(ln)
                and MONEY_RE.findall(ln)
            ),
            2,
        )

    for campo, padrao in LABELS_TXT.items():
        m = re.search(padrao, texto, re.I)
        dados[campo] = m.group(1).strip() if m else ""

    if not dados["nome"] and dados["cpf"]:
        m_nome = NOME_ANTES_CPF_RE.search(texto)
        if m_nome:
            dados["nome"] = m_nome.group(1).strip()

    return HoleriteAuditado(**dados)


def brl(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def mostrar_tabela(d: "pd.DataFrame"):
    """Compatível com versões novas e antigas do Streamlit."""
    try:
        st.dataframe(d, width="stretch")
    except TypeError:
        st.dataframe(d, use_container_width=True)


def responder(pergunta: str, d: "pd.DataFrame") -> str:
    """Respostas por regras sobre a tabela (não usa modelo de IA)."""
    p = pergunta.lower()
    if "diverg" in p:
        divs = d[d["divergencia"]]
        if divs.empty:
            return "Nenhuma divergência: todos os líquidos conferem com bruto − descontos."
        linhas = [
            f"- {r.nome or r.arquivo} (pág. {r.pagina}): "
            + ("valores não lidos (bruto zerado), revisar manualmente"
               if r.salario_bruto == 0 else f"diferença de {brl(r.diferenca)}")
            for r in divs.itertuples()
        ]
        return "Divergências encontradas:\n" + "\n".join(linhas)
    for chave, col, rot in [
        ("inss", "inss", "INSS"), ("irrf", "irrf", "IRRF"), ("fgts", "fgts", "FGTS"),
        ("líquid", "salario_liquido", "Líquido"), ("liquid", "salario_liquido", "Líquido"),
        ("bruto", "salario_bruto", "Bruto"), ("desconto", "total_descontos", "Descontos"),
    ]:
        if chave in p:
            return f"Total de {rot} no lote: {brl(d[col].sum())}."
    return (
        f"**Resumo do lote ({len(d)} folhas):** Bruto {brl(d['salario_bruto'].sum())} | "
        f"Líquido {brl(d['salario_liquido'].sum())} | INSS {brl(d['inss'].sum())} | "
        f"FGTS {brl(d['fgts'].sum())}. Divergências: {int(d['divergencia'].sum())}. "
        "Pergunte por: bruto, líquido, INSS, IRRF, FGTS, descontos ou divergências."
    )


st.title("Centralizador Seguro de Folhas de Pagamento (On-Premise)")
st.caption("Processamento 100% local · Conformidade LGPD · Sem envio de dados para nuvem externa.")

with st.sidebar:
    st.subheader("Governança & LGPD")
    mascarar = st.checkbox("Mascarar CPF na exibição e exportação", value=True)
    st.info(f"OCR local: {'Ativo (Tesseract)' if OCR_DISPONIVEL else 'Inativo (apenas PDF digital)'}")
    with st.expander("Diagnóstico"):
        import pydantic
        st.write(f"Python {platform.python_version()}")
        st.write(f"Streamlit {st.__version__}")
        st.write(f"pydantic {pydantic.VERSION}")
        st.write(f"pandas {pd.__version__}")

arquivos = st.file_uploader(
    "Selecione folhas de pagamento (PDF, PNG, JPG, TXT)",
    type=["pdf", "png", "jpg", "jpeg", "txt"],
    accept_multiple_files=True,
)

if arquivos and st.button("Processar e Auditar Lote", type="primary"):
    registros: List[dict] = []
    debug_textos: dict = {}

    with st.status("Executando pipeline contábil local...", expanded=True) as status:
        st.write("1. Validando arquivos e integridade de leitura...")
        for arq in arquivos:
            try:
                if arq.type == "application/pdf":
                    with pdfplumber.open(arq) as pdf:
                        for n, pag in enumerate(pdf.pages, start=1):
                            texto = extrair_texto_pagina(pag)
                            debug_textos[f"{arq.name} (pág. {n})"] = texto
                            registros.append(interpretar_local(arq.name, n, texto).model_dump())
                elif arq.type == "text/plain":
                    texto = arq.read().decode("utf-8", errors="ignore")
                    debug_textos[arq.name] = texto
                    registros.append(interpretar_local(arq.name, 1, texto).model_dump())
                elif OCR_DISPONIVEL:
                    texto = pytesseract.image_to_string(Image.open(arq), lang="por")
                    debug_textos[arq.name] = texto
                    registros.append(interpretar_local(arq.name, 1, texto).model_dump())
                else:
                    st.warning(f"{arq.name}: imagem ignorada porque Tesseract OCR não está instalado.")
            except Exception as e:
                st.error(f"Erro ao processar {arq.name}: {e}")
                st.exception(e)

        st.write("2. Executando auditoria aritmética (Bruto - Descontos = Líquido)...")
        df_resultado = pd.DataFrame(registros)
        if df_resultado.empty:
            st.warning("Nenhum registro foi extraído. Veja as mensagens acima.")
        st.session_state["df"] = df_resultado
        st.session_state["debug"] = debug_textos
        status.update(label=f"Lote concluído: {len(df_resultado)} registro(s) auditado(s).", state="complete")

df: Optional[pd.DataFrame] = st.session_state.get("df")
if df is not None and not df.empty:
    saida = df.copy()
    if mascarar:
        saida["cpf"] = saida["cpf"].map(mascarar_cpf)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Bruto", f"R$ {saida['salario_bruto'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c2.metric("Total Líquido", f"R$ {saida['salario_liquido'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c3.metric("INSS + IRRF", f"R$ {(saida['inss'].sum() + saida['irrf'].sum()):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c4.metric("Divergências", int(saida["divergencia"].sum()))

    mostrar_tabela(saida)
    try:
        graf = saida.copy()
        sem_nome = graf["nome"].astype(str).str.strip() == ""
        graf["rotulo"] = graf["nome"].astype(str).where(
            ~sem_nome, graf["arquivo"].astype(str) + " p" + graf["pagina"].astype(str))
        graf["rotulo"] = graf["rotulo"] + graf.groupby("rotulo").cumcount().map(
            lambda i: "" if i == 0 else f" ({i + 1})")
        st.bar_chart(graf.set_index("rotulo")[["salario_liquido", "inss", "irrf", "fgts"]])
    except Exception as e:
        st.caption(f"Gráfico indisponível: {e}")

    csv_bytes = saida.drop(columns=["divergencia", "diferenca"]).to_csv(
        index=False, sep=";", decimal=",", encoding="utf-8-sig"
    )
    st.download_button("Baixar CSV (Excel PT-BR)", csv_bytes, file_name="folhas_extraidas.csv", mime="text/csv")

    st.divider()
    st.subheader("Consulta rápida do lote (local)")
    st.caption("Responde com cálculos sobre a tabela acima; não usa modelo de IA nem internet.")
    if "msgs" not in st.session_state:
        st.session_state["msgs"] = []
    for msg in st.session_state["msgs"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("Pergunte sobre totais, INSS, FGTS ou divergências do lote..."):
        st.session_state["msgs"].append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        with st.chat_message("assistant"):
            resp = responder(prompt, saida)
            st.markdown(resp)
            st.session_state["msgs"].append({"role": "assistant", "content": resp})
