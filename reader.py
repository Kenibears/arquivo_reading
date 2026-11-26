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
import unicodedata
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


# ---------------------------------------------------------------------------
# Extração de NOME, CPF, CARGO e COMPETÊNCIA (vários formatos de holerite)
# ---------------------------------------------------------------------------
_MAI = "A-ZÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇÑ"
_MIN = "a-záàâãäéèêëíìîïóòôõöúùûüçñ"
_PALAVRA = rf"[{_MAI}][{_MAI}{_MIN}'.\-]+"
# 2+ palavras iniciadas por maiúscula (aceita "JOÃO DA SILVA" e "João da Silva")
NOME_RE = re.compile(rf"{_PALAVRA}(?:\s+(?:(?:d[aeo]s?|e|di|du|del|van|von)\s+)?{_PALAVRA})+")
CORTE_NOME_RE = re.compile(
    r"\s+(?:cpf\b|cargo\b|fun[cç][aã]o\b|ref\.?(?=[\s:]|$)|refer[eê]ncia\b|compet[eê]ncia\b|"
    r"cbo\b|admiss[aã]o\b|depart\w*|matr[ií]cula\b|c[oó]digo\b|sal[aá]rio\b)",
    re.I,
)
EMPRESA_RE = re.compile(
    r"\b(?:ltda|s\.?\s?a\.?|eireli|cnpj|empresa|com[eé]rcio|ind[uú]stria|servi[cç]os|"
    r"associa[cç][aã]o|funda[cç][aã]o|cooperativa)\b", re.I)
EVITA_RE = re.compile(
    r"cargo|fun[cç][aã]o|compet[eê]ncia|admiss|depart|\bref\b|\brg\b|\bpis\b|ctps|"
    r"recibo|demonstrativo|pagamento|holerite|folha", re.I)

TITULO_RE = re.compile(r"recibo|demonstrativo|pagamento|holerite|folha|declara", re.I)

_ROT_NOME = (
    r"nome(?!\s+d[aeo]s?\s+(?:m[aã]e|pai|banco|empresa|fantasia|social|dependente))"
    r"(?:\s+d[aeo]s?\s+(?:funcion[aá]ri[oa]|colaborador(?:a)?|empregad[oa]|trabalhador(?:a)?|servidor(?:a)?))?"
    r"|funcion[aá]ri[oa](?:\(a\))?|colaborador(?:\(a\)|a)?|empregad[oa]|trabalhador(?:a)?"
    r"|servidor(?:a)?|benefici[aá]ri[oa]"
)
ROTULO_NOME_RE = re.compile(rf"\b(?:{_ROT_NOME})(?!\w)", re.I)
SEP_RE = re.compile(r"^(?:[\s.]*[:\-–][\s.]*|\s*\.{3,}\s*)")  # ":"  "-"  "......"

CPF_ROTULO_RE = re.compile(r"cpf\b[^\d]{0,25}?(\d{3}\.?\d{3}\.?\d{3}-?\d{2})(?!\d)", re.I)
CPF_PONTUADO_RE = re.compile(r"(?<![\d./-])\d{3}\.\d{3}\.\d{3}-\d{2}(?![\d/])")

ROTULO_CARGO_RE = re.compile(r"\b(?:cargo|fun[cç][aã]o|ocupa[cç][aã]o)(?!\w)", re.I)
CORTE_CARGO_RE = re.compile(
    r"\s+(?:cpf\b|ref\.?(?=[\s:]|$)|compet[eê]ncia\b|cbo\b|admiss[aã]o\b|depart\w*|"
    r"matr[ií]cula\b|sal[aá]rio\b)|\s{2,}", re.I)

_MESES = {"janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
          "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}
COMP_ROTULO_RE = re.compile(
    r"(?:compet[eê]ncia|refer[eê]ncia|ref\.|m[eê]s\s*/\s*ano|per[ií]odo)\W{0,15}?"
    r"(0?[1-9]|1[0-2])\s*/\s*(\d{4})", re.I)
COMP_SOLTA_RE = re.compile(r"(?<![\d/])(0[1-9]|1[0-2])/((?:19|20)\d{2})(?!\d)")
COMP_MES_RE = re.compile(r"\b(" + "|".join(_MESES) + r")\s*(?:/|de|-)?\s*((?:19|20)\d{2})\b", re.I)


def limpar_nome(s: str) -> str:
    s = CORTE_NOME_RE.split(s, maxsplit=1)[0]
    s = re.sub(r"\s{2,}.*$", "", s)  # 2+ espaços normalmente separam colunas
    return s.strip(" .:-–\t")


def achar_cpf(texto: str):
    """Devolve (cpf formatado, posição no texto). Aceita 'CPF: 12345678909' sem pontuação."""
    m = CPF_ROTULO_RE.search(texto)
    if m:
        ini, bruto = m.start(1), m.group(1)
    else:
        m = CPF_PONTUADO_RE.search(texto)
        if not m:
            return "", -1
        ini, bruto = m.start(), m.group(0)
    d = re.sub(r"\D", "", bruto)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}", ini


def extrair_nome(texto: str, linhas: list, pos_cpf: int) -> str:
    # a) "Rótulo: valor" na mesma linha (Nome:, Funcionário -, Colaborador(a) ....:)
    for ln in linhas:
        for m in ROTULO_NOME_RE.finditer(ln):
            resto = ln[m.end():]
            sep = SEP_RE.match(resto)
            if not sep:
                continue
            cand = NOME_RE.match(resto[sep.end():].lstrip())
            if cand:
                nome = limpar_nome(cand.group(0))
                if nome and not EMPRESA_RE.search(nome):
                    return nome

    # b) rótulo em uma linha (cabeçalho de tabela) e o valor na linha de baixo
    for i, ln in enumerate(linhas[:-1]):
        if ":" in ln or TITULO_RE.search(ln) or not ROTULO_NOME_RE.search(ln):
            continue
        for prox in linhas[i + 1:i + 3]:
            if not prox.strip() or ROTULO_NOME_RE.search(prox):
                continue
            cand = NOME_RE.search(prox)
            if cand:
                nome = limpar_nome(cand.group(0))
                if nome and not EMPRESA_RE.search(nome):
                    return nome
            break

    # c) nome escrito logo antes do CPF (mesma linha ou linhas anteriores)
    if pos_cpf >= 0:
        ini_linha = texto.rfind("\n", 0, pos_cpf) + 1
        segmentos = [re.sub(r"(?i)cpf\b[\s:.\-]*$", "", texto[ini_linha:pos_cpf])]
        fim = ini_linha - 1
        for _ in range(3):
            if fim <= 0:
                break
            ini = texto.rfind("\n", 0, fim) + 1
            segmentos.append(texto[ini:fim])
            fim = ini - 1
        for seg in segmentos:
            if not seg.strip() or EMPRESA_RE.search(seg) or EVITA_RE.search(seg):
                continue
            achados = list(NOME_RE.finditer(seg))
            if achados:
                nome = limpar_nome(achados[0].group(0))
                if nome:
                    return nome
    return ""


def extrair_cargo(linhas: list) -> str:
    for ln in linhas:
        for m in ROTULO_CARGO_RE.finditer(ln):
            resto = ln[m.end():]
            sep = SEP_RE.match(resto)
            if not sep:
                continue
            valor = CORTE_CARGO_RE.split(resto[sep.end():].strip(), maxsplit=1)[0].strip(" .:-–")
            if valor and re.search(r"[A-Za-zÀ-ÿ]", valor):
                return valor
    return ""


def _sem_acento(s: str) -> str:
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()


def extrair_competencia(texto: str) -> str:
    m = COMP_ROTULO_RE.search(texto)
    if m:
        return f"{int(m.group(1)):02d}/{m.group(2)}"
    m = COMP_MES_RE.search(_sem_acento(texto))
    if m:
        return f"{_MESES[m.group(1).lower()]:02d}/{m.group(2)}"
    m = COMP_SOLTA_RE.search(texto)  # ignora datas completas tipo 05/01/2020
    return f"{m.group(1)}/{m.group(2)}" if m else ""


def cpf_valido(cpf: str) -> bool:
    d = [int(x) for x in re.sub(r"\D", "", cpf or "")]
    if len(d) != 11 or len(set(d)) == 1:
        return False
    for n in (9, 10):
        s = sum(v * w for v, w in zip(d[:n], range(n + 1, 1, -1)))
        if (s * 10 % 11) % 10 != d[n]:
            return False
    return True


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
    alertas: str = ""

    @model_validator(mode="after")
    def conferir_totais(self):
        esperado = round(self.salario_bruto - self.total_descontos, 2)
        self.diferenca = round(self.salario_liquido - esperado, 2)
        # Bruto zerado = documento não lido: também deve ser sinalizado para revisão.
        self.divergencia = abs(self.diferenca) > TOLERANCIA or self.salario_bruto == 0

        alertas = []
        if not self.nome:
            alertas.append("nome não encontrado")
        if not self.cpf:
            alertas.append("CPF não encontrado")
        elif not cpf_valido(self.cpf):
            alertas.append("CPF com dígito verificador inválido (conferir leitura)")
        if not self.competencia:
            alertas.append("competência não encontrada")
        if self.salario_bruto == 0:
            alertas.append("valores não lidos (bruto zerado)")
        elif self.divergencia:
            alertas.append("bruto − descontos ≠ líquido")
        self.alertas = "; ".join(alertas)
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

    dados["cpf"], pos_cpf = achar_cpf(texto)
    dados["nome"] = extrair_nome(texto, linhas, pos_cpf)
    dados["cargo"] = extrair_cargo(linhas)
    dados["competencia"] = extrair_competencia(texto)

    # Rede de segurança: padrões antigos (LABELS_TXT) só para o que ficou vazio
    for campo in ("nome", "cargo"):
        if not dados[campo]:
            m = re.search(LABELS_TXT[campo], texto, re.I)
            dados[campo] = limpar_nome(m.group(1)) if (m and campo == "nome") else (m.group(1).strip() if m else "")

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
    if any(k in p for k in ("alerta", "revis", "faltan", "sem nome", "sem cpf", "não encontr", "nao encontr")):
        com = d[d["alertas"] != ""]
        if com.empty:
            return "Nenhum alerta: nome, CPF, competência e totais foram lidos em todas as folhas."
        return "Folhas para revisar:\n" + "\n".join(
            f"- {r.nome or r.arquivo} (pág. {r.pagina}): {r.alertas}" for r in com.itertuples())
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

    csv_bytes = saida.drop(columns=["divergencia", "diferenca", "alertas"]).to_csv(
        index=False, sep=";", decimal=",", encoding="utf-8-sig"
    )
    st.download_button("Baixar CSV (Excel PT-BR)", csv_bytes, file_name="folhas_extraidas.csv", mime="text/csv")

    with st.expander("Texto extraído de cada página (para calibrar nome, CPF e demais campos)"):
        for nome_dbg, t in st.session_state.get("debug", {}).items():
            if mascarar:
                t = re.sub(r"(\d{3})\.(\d{3})\.(\d{3})-(\d{2})", r"•••.•••.•••-\4", t)
            st.markdown(f"**{nome_dbg}**")
            st.code(t[:3000] or "(vazio)")

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
