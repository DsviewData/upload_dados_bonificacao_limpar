import streamlit as st
import pandas as pd

# TESTE MÍNIMO - SEM AUTENTICAÇÃO
st.set_page_config(page_title="Teste Mínimo", layout="wide")

st.title("Teste Mínimo - Sistema de Bonificações")

st.success("✅ O app está funcionando!")

st.info("Se você está vendo esta mensagem, o Streamlit está OK")

# Testar secrets
st.subheader("Teste de Secrets")

try:
    client_id = st.secrets["CLIENT_ID"]
    st.success("✅ CLIENT_ID encontrado")
except:
    st.error("❌ CLIENT_ID não encontrado")

try:
    client_secret = st.secrets["CLIENT_SECRET"]
    st.success("✅ CLIENT_SECRET encontrado")
except:
    st.error("❌ CLIENT_SECRET não encontrado")

try:
    tenant_id = st.secrets["TENANT_ID"]
    st.success("✅ TENANT_ID encontrado")
except:
    st.error("❌ TENANT_ID não encontrado")

try:
    email = st.secrets["EMAIL_ONEDRIVE"]
    st.success("✅ EMAIL_ONEDRIVE encontrado")
except:
    st.error("❌ EMAIL_ONEDRIVE não encontrado")

try:
    site_id = st.secrets["SITE_ID"]
    st.success("✅ SITE_ID encontrado")
except:
    st.error("❌ SITE_ID não encontrado")

try:
    drive_id = st.secrets["DRIVE_ID"]
    st.success("✅ DRIVE_ID encontrado")
except:
    st.error("❌ DRIVE_ID não encontrado")

st.divider()

# Teste de upload simples
st.subheader("Teste de Upload")
uploaded_file = st.file_uploader("Teste de upload de arquivo", type=["xlsx", "xls"])

if uploaded_file:
    st.success(f"Arquivo carregado: {uploaded_file.name}")
    
    try:
        df = pd.read_excel(uploaded_file)
        st.dataframe(df.head())
        st.success(f"✅ Arquivo lido com sucesso: {len(df)} linhas")
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")

st.divider()
st.info("Se todas as secrets mostrarem ✅, o problema não é de configuração básica")
