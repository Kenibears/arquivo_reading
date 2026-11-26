import io
import re
from typing import List, Optional

import pandas as pd
import pdfplumber
import streamlit as st
from pydantic import BaseModel, Field, model_validator

try:
  import pytesseract
  from PIL import Image

  OCR_DISPONIVEL = True
except ImportError:
  OCR_DISPONIVEL = False

TOLERANCIA = 0.01
MONEY_RE = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}")

LABELS_NUM = {
    "salario_bruto": (
        r"sal[aá]rio\s+bruto|total\s+(?:de\s+)?(?:proventos|vencimentos)|total\s+bruto|sal[aá]rio\s+base(?:\s+mensal)?|sal\.?\s*base(?:\s+mensal)?|sal[aá]rio\s+mensal|sal[aá]rio\s+contratual|vencimento\s+base"
    ),
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

QUADRO_FGTS_RE = re.compile(
    r"fgts\s+do\s+m[eê]s\s*\(8\%\)[\s\S]*?(\d{1,3}(?:\.\d{3})*,\d{2})\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s+(\d{1,3}(?:\.\d{3})*,\d{2})",
    re.I,
)

TABELA_FUNC_RE = re.compile(
    r"\b\d{3,6}\s+([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ]{2,}(?:\s+[A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ.]{1,}){1,7})\s+(\d{3}\.\d{3}\.\d{3}-\d{2})\s+(.+?)\s+\d{2}/\d{2}/\d{4}"
)

NOME_ANTES_CPF_RE = re.compile(
    r"([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ]{2,}(?:[ \t]+[A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ.]{1,}){1,7})\s*(?:CPF\s*[:\-]?\s*)?\d{3}\.\d{3}\.\d{3}-\d{2}"
)

LABELS_TXT = {
    "nome": (
        r"(?m)nome(?:\s+do\s+funcion[aá]rio)?\s*[:\-]\s*(.+?)(?=\s+(?:ref\.?|cpf|cargo|compet[eê]ncia)\b|$)"
    ),
    "cpf": r"(\d{3}\.\d{3}\.\d{3}-\d{2})",
    "cargo": (
        r"(?m)cargo\s*[:\-]\s*(.+?)(?=\s+(?:ref\.?|cpf|compet[eê]ncia)\b|$)"
    ),
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
    self.divergencia = abs(self.diferenca) > TOLERANCIA
    return self


def br_para_float(valor: str) -> float:
  return float(valor.replace(".", "").replace(",", "."))


def mascarar_cpf(cpf: str) -> str:
  return re.sub(
      r"\d{3}\.\d{3}\.\d{3}-(\d{2})", r"•••.•••.•••-\1", cpf or ""
  )


def extrair_texto_pagina(pagina) -> str:
  txt = pagina.extract_text() or ""
  if len(txt.strip()) < 30 and OCR_DISPONIVEL:
    img = pagina.to_image(resolution=300).original
    txt = pytesseract.image_to_string(img, lang="por")
  return txt


def _extrair_valor_linha_ou_proxima(
    linhas: list[str], padrao: str | re.Pattern
) -> Optional[float]:
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


def interpretar_local(
    arquivo: str, pagina: int, texto: str
) -> HoleriteAuditado:
  dados: dict = {"arquivo": arquivo, "pagina": pagina}
  linhas = texto.splitlines()
  for campo in LABELS_NUM:
    dados[campo] = 0.0

  m_totais = TRES_TOTAIS_RE.search(texto)
  if m_totais:
    dados["salario_bruto"] = br_para_float(m_totais.group(1))
    dados["total_descontos"] = br_para_float(m_totais.group(2))
    dados["salario_liquido"] = br_para_float(m_totais.group(3))

  m_fgts = QUADRO_FGTS_RE.search(texto)
  if m_fgts:
    dados["fgts"] = br_para_float(m_fgts.group(4))

  for campo, padrao in LABELS_NUM.items():
    if dados[campo] > 0:
      continue
    if campo == "salario_bruto":
      prioritario = _extrair_valor_linha_ou_proxima(
          linhas, SALARIO_BRUTO_PRIORITARIO
      )
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

  m_tab = TABELA_FUNC_RE.search(texto)
  if m_tab:
    if not dados["nome"]:
      dados["nome"] = m_tab.group(1).strip()
    if not dados["cpf"]:
      dados["cpf"] = m_tab.group(2).strip()
    if not dados["cargo"]:
      dados["cargo"] = m_tab.group(3).strip()
  elif not dados["nome"] and dados["cpf"]:
    m_nome = NOME_ANTES_CPF_RE.search(texto)
    if m_nome:
      dados["nome"] = m_nome.group(1).strip()

  if not dados["nome"]:
    dados["nome"] = f"Colaborador (Pág. {pagina})"

  return HoleriteAuditado(**dados)


# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="ContábilFlow Local", layout="wide")
st.title("Centralizador Seguro de Folhas de Pagamento (On-Premise)")
st.caption(
    "Processamento 100% local · Conformidade LGPD · Sem envio de dados para"
    " nuvem externa."
)

with st.sidebar:
  st.subheader("Governança & LGPD")
  mascarar = st.checkbox("Mascarar CPF na exibição e exportação", value=True)
  st.info(
      f"OCR local: {'Ativo (Tesseract)' if OCR_DISPONIVEL else 'Inativo (apenas PDF digital)'}"
  )

arquivos = st.file_uploader(
    "Selecione folhas de pagamento (PDF, PNG, JPG, TXT)",
    type=["pdf", "png", "jpg", "jpeg", "txt"],
    accept_multiple_files=True,
)

if arquivos and st.button("Processar e Auditar Lote", type="primary"):
  registros: List[dict] = []
  debug_textos: dict = {}

  with st.status(
      "Executando pipeline contábil local...", expanded=True
  ) as status:
    st.write("1. Validando arquivos e integridade de leitura...")
    for arq in arquivos:
      if arq.type == "application/pdf":
        with pdfplumber.open(arq) as pdf:
          for n, pag in enumerate(pdf.pages, start=1):
            texto = extrair_texto_pagina(pag)
            debug_textos[f"{arq.name} (pág. {n})"] = texto
            registros.append(
                interpretar_local(arq.name, n, texto).model_dump()
            )
      elif arq.type == "text/plain":
        texto = arq.read().decode("utf-8", errors="ignore")
        debug_textos[arq.name] = texto
        registros.append(interpretar_local(arq.name, 1, texto).model_dump())
      elif OCR_DISPONIVEL:
        texto = pytesseract.image_to_string(Image.open(arq), lang="por")
        debug_textos[arq.name] = texto
        registros.append(interpretar_local(arq.name, 1, texto).model_dump())
      else:
        st.warning(
            f"{arq.name}: imagem ignorada porque Tesseract OCR não está"
            " instalado."
        )

    st.write(
        "2. Executando auditoria aritmética (Bruto - Descontos = Líquido)..."
    )
    df_resultado = pd.DataFrame(registros)
    st.session_state["df"] = df_resultado
    st.session_state["debug"] = debug_textos
    status.update(
        label=f"Lote concluído: {len(df_resultado)} registro(s) auditado(s).",
        state="complete",
    )

df: Optional[pd.DataFrame] = st.session_state.get("df")
if df is not None and not df.empty:
  saida = df.copy()
  if mascarar:
    saida["cpf"] = saida["cpf"].map(mascarar_cpf)

  c1, c2, c3, c4 = st.columns(4)
  c1.metric(
      "Total Bruto",
      f"R$ {saida['salario_bruto'].sum():,.2f}".replace(",", "X")
      .replace(".", ",")
      .replace("X", "."),
  )
  c2.metric(
      "Total Líquido",
      f"R$ {saida['salario_liquido'].sum():,.2f}".replace(",", "X")
      .replace(".", ",")
      .replace("X", "."),
  )
  c3.metric(
      "INSS + IRRF
