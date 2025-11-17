# app.py

import streamlit as st
from process import destacar_nups

st.set_page_config(page_title="Destacar NUPs Duplicados", page_icon="🔍")

st.title("🔍 Destacar NUPs Duplicados")
st.write("""
Envie um arquivo Excel (.xlsx) e o sistema irá identificar automaticamente 
os **NUPs duplicados**, destacando-os em **cores alternadas** para facilitar a visualização.
""")

uploaded = st.file_uploader("📂 Envie um arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded:
    st.success("Arquivo carregado com sucesso! Processando...")

    result_file = destacar_nups(uploaded.getvalue())

    st.download_button(
        label="📥 Baixar arquivo com NUPs destacados",
        data=result_file,
        file_name=uploaded.name.replace(".xlsx", "_destacado.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.info("Clique no botão acima para baixar o arquivo processado.")
