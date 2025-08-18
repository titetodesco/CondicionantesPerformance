import streamlit as st
import pandas as pd
import plotly.express as px
from rapidfuzz import fuzz
import fitz  # PyMuPDF
from docx import Document
from unidecode import unidecode
import os
import tempfile

st.set_page_config(layout="wide")
st.title("🛠️ Análise de Acidentes – Versão GitHub")

# Caminho fixo da planilha no repositório
TAXONOMIA_PATH = "TaxonomiaCP_Por.xlsx"
TAXONOMIA_PATH = "https://raw.githubusercontent.com/titetodesco/CondicionantesPerformance/main/TaxonomiaCP_Por.xlsx"

# Sidebar – upload do relatório
st.sidebar.header("📂 Upload do Relatório de Acidente")
file_relato = st.sidebar.file_uploader(
    "Relatório (.pdf, .docx, .txt)",
    type=["pdf", "docx", "txt"]
)

# Função para extrair texto
def extract_text(file):
    suffix = os.path.splitext(file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file.read())
        tmp_path = tmp.name

    if suffix == ".pdf":
        doc = fitz.open(tmp_path)
        text = "\n".join([page.get_text() for page in doc])
    elif suffix == ".docx":
        doc = Document(tmp_path)
        text = "\n".join([para.text for para in doc.paragraphs])
    elif suffix == ".txt":
        with open(tmp_path, "r", encoding="utf-8") as f:
            text = f.read()
    else:
        text = ""

    return unidecode(text.lower())

# Função para detectar fatores
def detectar_fatores(texto, df, coluna_termos):
    resultados = []
    for _, row in df.iterrows():
        termos = row.get(coluna_termos, [])
        for termo in termos:
            if termo in texto or fuzz.partial_ratio(termo, texto) > 85:
                resultados.append({
                    "Dimensão": row["Dimensão"],
                    "Fatores": row["Fatores"],
                    "Subfator 1": row.get("Subfator 1"),
                    "Subfator 2": row.get("Subfator 2"),
                    "Recomendação 1": row.get("Recomendação 1", ""),
                    "Recomendação 2": row.get("Recomendação 2", ""),
                    "Termo identificado": termo,
                    "Similaridade": 100.0 if termo in texto else fuzz.partial_ratio(termo, texto)
                })
    return pd.DataFrame(resultados)

# Fluxo principal
if file_relato:
    # Carrega texto do relatório
    texto = extract_text(file_relato)

    # Carrega planilha fixa
    if not os.path.exists(TAXONOMIA_PATH):
        st.error(f"Planilha {TAXONOMIA_PATH} não encontrada no repositório.")
        st.stop()

    df_cp = pd.read_excel(TAXONOMIA_PATH)

    # Detecta idioma
    idioma = "pt" if any(p in texto for p in ["segurança", "procedimento", "acidente", "falha", "trabalho"]) else "en"
    coluna_termos = "Bag de termos" if idioma == "pt" else "Bag of terms"
    st.info(f"🌍 Idioma detectado: {'Português' if idioma == 'pt' else 'Inglês'} – usando `{coluna_termos}`")

    # Prepara termos
    df_cp[coluna_termos] = df_cp[coluna_termos].fillna("").apply(
        lambda x: [t.strip().lower() for t in str(x).split(";") if t.strip()]
    )

    # Detecta fatores
    resultados = detectar_fatores(texto, df_cp, coluna_termos)

    if resultados.empty:
        st.warning("⚠ Nenhum termo identificado.")
    else:
        st.success("✅ Termos identificados!")

        aba = st.selectbox(
            "Escolha a visualização:",
            ["📌 Resumo por Fator", "📊 Gráficos", "🧠 Recomendações"]
        )

        if aba == "📌 Resumo por Fator":
            resumo_dimensao = resultados["Dimensão"].value_counts().reset_index()
            resumo_dimensao.columns = ["Dimensão", "Frequência"]
            st.dataframe(resumo_dimensao)

        elif aba == "📊 Gráficos":
            fig1 = px.histogram(resultados, y="Dimensão", color="Dimensão", title="Frequência por Dimensão", text_auto=True)
            st.plotly_chart(fig1, use_container_width=True)

        elif aba == "🧠 Recomendações":
            df_rec = resultados.groupby(
                ["Dimensão", "Fatores", "Recomendação 1", "Recomendação 2"]
            ).size().reset_index(name="Frequência")
            st.dataframe(df_rec)
else:
    st.info("📤 Faça upload de um relatório para iniciar.")
