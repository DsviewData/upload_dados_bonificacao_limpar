import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
from io import BytesIO
from msal import ConfidentialClientApplication
import logging
import json
import uuid
import time

# ===========================
# CONFIGURAÇÕES DE VERSÃO
# ===========================
APP_VERSION = "1.0.1"
VERSION_DATE = "2025-10-03"

# ===========================
# CONFIGURAÇÃO DE LOGGING
# ===========================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ===========================
# CREDENCIAIS VIA ST.SECRETS - COM VERIFICAÇÃO
# ===========================
CREDENCIAIS_OK = False
CREDENCIAL_FALTANDO = ""

try:
    CLIENT_ID = st.secrets["CLIENT_ID"]
    CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
    TENANT_ID = st.secrets["TENANT_ID"]
    EMAIL_ONEDRIVE = st.secrets["EMAIL_ONEDRIVE"]
    SITE_ID = st.secrets["SITE_ID"]
    DRIVE_ID = st.secrets["DRIVE_ID"]
    CREDENCIAIS_OK = True
except KeyError as e:
    CREDENCIAL_FALTANDO = str(e)

# ===========================
# CONFIGURAÇÃO DE PASTAS
# ===========================
PASTA_CONSOLIDADO = "Documentos Compartilhados/Bonificacao/FonteDeDados"
PASTA_ENVIOS_BACKUPS = "Documentos Compartilhados/PlanilhasEnviadas_Backups/Bonificacao"
ARQUIVO_LOCK = "sistema_lock_bonificacao.json"
TIMEOUT_LOCK_MINUTOS = 10

# ===========================
# ESTILOS CSS
# ===========================
def aplicar_estilos_css():
    st.markdown("""
    <style>
    :root {
        --primary-color: #2E8B57;
        --secondary-color: #20B2AA;
        --success-color: #32CD32;
        --warning-color: #FFA500;
        --error-color: #DC143C;
    }
    
    .main-header {
        background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        color: white;
    }
    
    .main-header h1 {
        margin: 0;
        font-size: 2rem;
        font-weight: 700;
    }
    
    .status-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border-left: 4px solid var(--primary-color);
        margin: 1rem 0;
    }
    
    .status-card.success {
        border-left-color: var(--success-color);
        background: linear-gradient(135deg, #f0fff4, #ffffff);
    }
    
    .status-card.error {
        border-left-color: var(--error-color);
        background: linear-gradient(135deg, #fff0f0, #ffffff);
    }
    
    .metric-container {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border-top: 3px solid var(--primary-color);
    }
    
    .metric-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: var(--primary-color);
        margin: 0.5rem 0;
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: #2C3E50;
        font-weight: 500;
        text-transform: uppercase;
    }
    </style>
    """, unsafe_allow_html=True)

# ===========================
# AUTENTICAÇÃO - SEM CACHE PROBLEMÁTICO
# ===========================
def obter_token():
    """Obtém token de acesso - usa session_state ao invés de cache"""
    if 'token_cache' not in st.session_state or \
       'token_timestamp' not in st.session_state or \
       (datetime.now() - st.session_state.token_timestamp).seconds > 3000:
        
        try:
            app = ConfidentialClientApplication(
                CLIENT_ID,
                authority=f"https://login.microsoftonline.com/{TENANT_ID}",
                client_credential=CLIENT_SECRET
            )
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            
            if "access_token" not in result:
                error_desc = result.get("error_description", "Token não obtido")
                logger.error(f"Falha na autenticação: {error_desc}")
                return None
            
            st.session_state.token_cache = result["access_token"]
            st.session_state.token_timestamp = datetime.now()
            logger.info("Token obtido com sucesso")
            return result["access_token"]
            
        except Exception as e:
            logger.error(f"Erro de autenticação: {e}")
            return None
    
    return st.session_state.token_cache

# ===========================
# SISTEMA DE LOCK
# ===========================
def gerar_id_sessao():
    if 'session_id' not in st.session_state:
        st.session_state.session_id = str(uuid.uuid4())[:8]
    return st.session_state.session_id

def verificar_lock_existente(token):
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{ARQUIVO_LOCK}:/content"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            lock_data = response.json()
            timestamp_lock = datetime.fromisoformat(lock_data['timestamp'])
            agora = datetime.now()
            
            if agora - timestamp_lock > timedelta(minutes=TIMEOUT_LOCK_MINUTOS):
                logger.info("Lock expirado - removendo")
                remover_lock(token, force=True)
                return False, None
            
            return True, lock_data
        
        return False, None
            
    except Exception as e:
        logger.error(f"Erro ao verificar lock: {e}")
        return False, None

def criar_lock(token, operacao="Consolidação de bonificações"):
    try:
        session_id = gerar_id_sessao()
        
        lock_data = {
            "timestamp": datetime.now().isoformat(),
            "session_id": session_id,
            "operacao": operacao,
            "status": "EM_ANDAMENTO",
            "app_version": APP_VERSION
        }
        
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{ARQUIVO_LOCK}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        response = requests.put(url, headers=headers, data=json.dumps(lock_data), timeout=10)
        
        if response.status_code in [200, 201]:
            logger.info(f"Lock criado: {session_id}")
            return True, session_id
        return False, None
            
    except Exception as e:
        logger.error(f"Erro ao criar lock: {e}")
        return False, None

def remover_lock(token, session_id=None, force=False):
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{ARQUIVO_LOCK}"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.delete(url, headers=headers, timeout=10)
        
        if response.status_code in [200, 204, 404]:
            logger.info("Lock removido")
            return True
        return False
            
    except Exception as e:
        logger.error(f"Erro ao remover lock: {e}")
        return False

def exibir_status_sistema(token):
    lock_existe, lock_data = verificar_lock_existente(token)
    
    if lock_existe:
        timestamp_inicio = datetime.fromisoformat(lock_data['timestamp'])
        duracao = datetime.now() - timestamp_inicio
        
        st.markdown("""
        <div class="status-card error">
            <h3>Sistema Ocupado</h3>
            <p>Outro usuário está enviando dados no momento</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Tempo Ativo", f"{int(duracao.total_seconds()//60)} min")
        with col2:
            st.metric("Operação", lock_data.get('operacao', 'N/A'))
        
        return True
    else:
        st.markdown("""
        <div class="status-card success">
            <h3>Sistema Disponível</h3>
            <p>Você pode enviar sua planilha agora</p>
        </div>
        """, unsafe_allow_html=True)
        return False

# ===========================
# FUNÇÕES AUXILIARES
# ===========================
def criar_pasta_se_nao_existir(caminho_pasta, token):
    try:
        partes = caminho_pasta.split('/')
        caminho_atual = ""
        
        for parte in partes:
            if not parte:
                continue
                
            caminho_anterior = caminho_atual
            caminho_atual = f"{caminho_atual}/{parte}" if caminho_atual else parte
            
            url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{caminho_atual}"
            headers = {"Authorization": f"Bearer {token}"}
            response = requests.get(url, headers=headers, timeout=10)
            
            if response.status_code == 404:
                parent_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root"
                if caminho_anterior:
                    parent_url += f":/{caminho_anterior}"
                parent_url += ":/children"
                
                create_body = {
                    "name": parte,
                    "folder": {},
                    "@microsoft.graph.conflictBehavior": "rename"
                }
                
                requests.post(
                    parent_url, 
                    headers={**headers, "Content-Type": "application/json"}, 
                    json=create_body,
                    timeout=10
                )
                    
    except Exception as e:
        logger.warning(f"Erro ao criar pastas: {e}")

def upload_onedrive(nome_arquivo, conteudo_arquivo, token, tipo_arquivo="consolidado"):
    try:
        if tipo_arquivo == "consolidado":
            pasta_base = PASTA_CONSOLIDADO
        else:
            pasta_base = PASTA_ENVIOS_BACKUPS
        
        criar_pasta_se_nao_existir(pasta_base, token)
        
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{pasta_base}/{nome_arquivo}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }
        response = requests.put(url, headers=headers, data=conteudo_arquivo, timeout=60)
        
        return response.status_code in [200, 201]
        
    except Exception as e:
        logger.error(f"Erro no upload: {e}")
        return False

def baixar_arquivo_consolidado(token):
    consolidado_nome = "bonificacao_consolidada.xlsx"
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{consolidado_nome}:/content"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            df_consolidado = pd.read_excel(BytesIO(response.content))
            df_consolidado.columns = df_consolidado.columns.str.strip().str.upper()
            logger.info(f"Arquivo consolidado baixado: {len(df_consolidado)} registros")
            return df_consolidado, True
        else:
            logger.info("Arquivo consolidado não existe - será criado novo")
            return pd.DataFrame(), False
            
    except Exception as e:
        logger.error(f"Erro ao baixar: {e}")
        return pd.DataFrame(), False

def salvar_arquivo_enviado(df_novo, nome_arquivo_original, token):
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d_%Hh%M")
        nome_base = nome_arquivo_original.replace(".xlsx", "").replace(".xls", "")
        nome_arquivo_backup = f"{nome_base}_enviado_{timestamp}.xlsx"
        
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_novo.to_excel(writer, index=False, sheet_name="Dados")
        buffer.seek(0)
        
        upload_onedrive(nome_arquivo_backup, buffer.read(), token, "backup")
        logger.info(f"Backup salvo: {nome_arquivo_backup}")
            
    except Exception as e:
        logger.error(f"Erro ao salvar backup: {e}")

# ===========================
# VALIDAÇÃO
# ===========================
def validar_dados_enviados(df):
    erros = []
    avisos = []
    
    if df.empty:
        erros.append("A planilha está vazia")
        return erros, avisos
    
    if "LOJA" not in df.columns:
        erros.append("A planilha deve conter uma coluna 'LOJA'")
    else:
        lojas_validas = df["LOJA"].notna().sum()
        if lojas_validas == 0:
            erros.append("Nenhuma loja válida encontrada")
        else:
            lojas_unicas = df["LOJA"].dropna().unique()
            avisos.append(f"Lojas encontradas: {len(lojas_unicas)}")
    
    if "DATA" not in df.columns:
        erros.append("A planilha deve conter uma coluna 'DATA'")
    else:
        try:
            df_temp = df.copy()
            df_temp["DATA"] = pd.to_datetime(df_temp["DATA"], errors="coerce")
            datas_invalidas = df_temp["DATA"].isna().sum()
            if datas_invalidas > 0:
                avisos.append(f"{datas_invalidas} datas inválidas serão removidas")
        except:
            erros.append("Erro ao processar datas")
    
    return erros, avisos

# ===========================
# CONSOLIDAÇÃO
# ===========================
def adicionar_data_ultimo_envio(df_final, lojas_atualizadas):
    try:
        if 'DATA_ULTIMO_ENVIO' not in df_final.columns:
            df_final['DATA_ULTIMO_ENVIO'] = pd.NaT
        
        data_atual = datetime.now()
        
        for loja in lojas_atualizadas:
            mask = df_final['LOJA'].astype(str).str.strip().str.upper() == str(loja).strip().upper()
            df_final.loc[mask, 'DATA_ULTIMO_ENVIO'] = data_atual
        
        return df_final
        
    except Exception as e:
        logger.error(f"Erro ao adicionar data envio: {e}")
        return df_final

def comparar_e_atualizar_registros(df_consolidado, df_novo):
    registros_inseridos = 0
    registros_substituidos = 0
    registros_removidos = 0
    lojas_atualizadas = set()
    
    logger.info(f"Consolidação: {len(df_consolidado)} existentes + {len(df_novo)} novos")
    
    if df_consolidado.empty:
        df_final = df_novo.copy()
        registros_inseridos = len(df_novo)
        lojas_atualizadas = set(df_novo['LOJA'].dropna().astype(str).str.strip().str.upper().unique())
        df_final = adicionar_data_ultimo_envio(df_final, lojas_atualizadas)
        return df_final, registros_inseridos, registros_substituidos, registros_removidos
    
    # Garantir colunas
    colunas = df_novo.columns.tolist()
    for col in colunas:
        if col not in df_consolidado.columns:
            df_consolidado[col] = None
    
    df_final = df_consolidado.copy()
    
    # Adicionar mes_ano
    df_novo_temp = df_novo.copy()
    df_novo_temp['mes_ano'] = df_novo_temp['DATA'].dt.to_period('M')
    
    df_final_temp = df_final.copy()
    df_final_temp['mes_ano'] = df_final_temp['DATA'].dt.to_period('M')
    
    # Processar por loja + mes/ano
    grupos_novos = df_novo_temp.groupby(['LOJA', 'mes_ano'])
    
    for (loja, periodo_grupo), grupo_df in grupos_novos:
        if pd.isna(loja) or str(loja).strip() == '':
            continue
        
        lojas_atualizadas.add(str(loja).strip().upper())
        
        mask_existente = (
            (df_final_temp["mes_ano"] == periodo_grupo) &
            (df_final_temp["LOJA"].astype(str).str.strip().str.upper() == str(loja).strip().upper())
        )
        
        registros_existentes = df_final[mask_existente]
        
        if not registros_existentes.empty:
            num_removidos = len(registros_existentes)
            df_final = df_final[~mask_existente]
            df_final_temp = df_final_temp[~mask_existente]
            registros_removidos += num_removidos
            registros_substituidos += len(grupo_df)
        else:
            registros_inseridos += len(grupo_df)
        
        grupo_para_inserir = grupo_df.drop(columns=['mes_ano'], errors='ignore')
        df_final = pd.concat([df_final, grupo_para_inserir], ignore_index=True)
        df_final_temp = pd.concat([df_final_temp, grupo_df], ignore_index=True)
    
    df_final = adicionar_data_ultimo_envio(df_final, lojas_atualizadas)
    
    logger.info(f"Resultado: {registros_inseridos} inseridos, {registros_substituidos} substituídos")
    
    return df_final, registros_inseridos, registros_substituidos, registros_removidos

def analise_pre_consolidacao(df_consolidado, df_novo):
    try:
        st.subheader("Análise Pré-Consolidação")
        
        df_novo_temp = df_novo.copy()
        df_novo_temp['mes_ano'] = df_novo_temp['DATA'].dt.to_period('M')
        
        lojas_novas = set(df_novo['LOJA'].dropna().astype(str).str.strip().str.upper().unique())
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Lojas", len(lojas_novas))
        
        with col2:
            st.metric("Registros Novos", len(df_novo))
        
        with col3:
            periodos_unicos = df_novo_temp.groupby(['LOJA', 'mes_ano']).size()
            st.metric("Períodos", len(periodos_unicos))
        
        return True
        
    except Exception as e:
        logger.error(f"Erro na análise: {e}")
        st.error(f"Erro na análise: {str(e)}")
        return False

def processar_consolidacao_com_lock(df_novo, nome_arquivo, token):
    session_id = gerar_id_sessao()
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        status_text.info("Iniciando consolidação...")
        progress_bar.progress(10)
        
        # Verificar lock
        sistema_ocupado, _ = verificar_lock_existente(token)
        if sistema_ocupado:
            status_text.error("Sistema ocupado por outro usuário")
            return False
        
        # Criar lock
        status_text.info("Bloqueando sistema...")
        progress_bar.progress(15)
        
        lock_criado, session_lock = criar_lock(token)
        if not lock_criado:
            status_text.error("Não foi possível bloquear o sistema")
            return False
        
        # Baixar consolidado
        status_text.info("Baixando arquivo consolidado...")
        progress_bar.progress(30)
        
        df_consolidado, arquivo_existe = baixar_arquivo_consolidado(token)
        
        # Preparar dados
        status_text.info("Preparando dados...")
        progress_bar.progress(40)
        
        df_novo = df_novo.copy()
        df_novo.columns = df_novo.columns.str.strip().str.upper()
        df_novo["DATA"] = pd.to_datetime(df_novo["DATA"], errors="coerce")
        df_novo = df_novo.dropna(subset=["DATA"])
        
        if df_novo.empty:
            status_text.error("Nenhum registro válido")
            remover_lock(token, session_lock)
            return False
        
        # Análise
        status_text.info("Analisando dados...")
        progress_bar.progress(50)
        analise_pre_consolidacao(df_consolidado, df_novo)
        
        # Consolidar
        status_text.info("Consolidando...")
        progress_bar.progress(70)
        
        df_final, inseridos, substituidos, removidos = comparar_e_atualizar_registros(
            df_consolidado, df_novo
        )
        
        df_final = df_final.sort_values(["DATA", "LOJA"], na_position='last').reset_index(drop=True)
        
        # Salvar backup
        status_text.info("Salvando backup...")
        progress_bar.progress(85)
        salvar_arquivo_enviado(df_novo, nome_arquivo, token)
        
        # Upload final
        status_text.info("Salvando consolidado...")
        progress_bar.progress(90)
        
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name="Dados")
        buffer.seek(0)
        
        sucesso = upload_onedrive("bonificacao_consolidada.xlsx", buffer.read(), token)
        
        # Remover lock
        remover_lock(token, session_lock)
        progress_bar.progress(100)
        
        if sucesso:
            status_text.empty()
            progress_bar.empty()
            
            st.success("Consolidação realizada com sucesso!")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Final", f"{len(df_final):,}")
            with col2:
                st.metric("Inseridos", inseridos)
            with col3:
                st.metric("Substituídos", substituidos)
            with col4:
                st.metric("Removidos", removidos)
            
            if not df_final.empty and 'LOJA' in df_final.columns:
                resumo = df_final.groupby("LOJA").agg({
                    "DATA": ["count", "min", "max"]
                })
                resumo.columns = ["Total", "Data Inicial", "Data Final"]
                
                with st.expander("Resumo por Loja"):
                    st.dataframe(resumo, use_container_width=True)
            
            return True
        else:
            status_text.error("Erro ao salvar arquivo")
            return False
            
    except Exception as e:
        logger.error(f"Erro na consolidação: {e}")
        remover_lock(token, session_id, force=True)
        status_text.error(f"Erro: {str(e)}")
        progress_bar.empty()
        return False

# ===========================
# INTERFACE PRINCIPAL
# ===========================
def main():
    st.set_page_config(
        page_title=f"Bonificações v{APP_VERSION}", 
        layout="wide",
        initial_sidebar_state="expanded"
    )

    aplicar_estilos_css()

    st.markdown(f"""
    <div class="main-header">
        <h1>Sistema de Consolidação de Bonificações</h1>
        <p style="margin: 0.5rem 0 0 0; opacity: 0.9;">Upload e consolidação automática por Loja + Mês/Ano</p>
        <small>Versão {APP_VERSION} - {VERSION_DATE}</small>
    </div>
    """, unsafe_allow_html=True)

    # Verificação de credenciais
    if not CREDENCIAIS_OK:
        st.error(f"ERRO: Credencial não configurada: {CREDENCIAL_FALTANDO}")
        st.info("Configure todas as secrets no Streamlit Cloud:")
        st.code("""CLIENT_ID
CLIENT_SECRET
TENANT_ID
EMAIL_ONEDRIVE
SITE_ID
DRIVE_ID""")
        st.info("Vá em: Manage app > Settings > Secrets")
        st.stop()

    st.sidebar.markdown("### Upload de Bonificações")
    st.sidebar.divider()

    # Obter token
    with st.spinner("Conectando ao Microsoft Graph..."):
        token = obter_token()
    
    if not token:
        st.error("Erro de autenticação. Verifique as credenciais nas secrets.")
        st.sidebar.error("Desconectado")
        st.stop()
    
    st.sidebar.success("Conectado")

    # Status do sistema
    st.markdown("## Status do Sistema")
    sistema_ocupado = exibir_status_sistema(token)
    
    if sistema_ocupado:
        st.divider()
        if st.button("Atualizar Status"):
            st.rerun()
        st.info("Página será atualizada automaticamente em 15 segundos")
        time.sleep(15)
        st.rerun()

    st.divider()

    # Info sobre o sistema
    with st.sidebar.expander("Informações"):
        st.markdown(f"**Versão:** {APP_VERSION}")
        st.markdown(f"**Data:** {VERSION_DATE}")
        st.markdown("**Consolidado:** bonificacao_consolidada.xlsx")

    # Upload
    st.markdown("## Upload de Planilha Excel")
    
    st.info("A planilha deve ter uma aba 'Dados' com as colunas 'LOJA' e 'DATA'")
    
    with st.expander("Como funciona a consolidação", expanded=False):
        st.markdown("""
        **Consolidação por LOJA + MÊS/ANO:**
        - Substitui dados mensais existentes da mesma loja
        - Adiciona novos períodos mensais
        - Mantém dados de outras lojas intactos
        - Registra data do último envio por loja
        - Cria backups automáticos
        """)

    uploaded_file = st.file_uploader(
        "Escolha um arquivo Excel", 
        type=["xlsx", "xls"],
        help="Formatos aceitos: .xlsx, .xls"
    )

    df = None
    if uploaded_file:
        try:
            st.success(f"Arquivo carregado: {uploaded_file.name}")
            
            with st.spinner("Lendo arquivo..."):
                xls = pd.ExcelFile(uploaded_file)
                sheets = xls.sheet_names
                
                if "Dados" in sheets:
                    sheet = "Dados"
                    st.success("Aba 'Dados' encontrada automaticamente")
                else:
                    sheet = st.selectbox("Selecione a aba:", sheets)
                    if sheet != "Dados":
                        st.warning("Recomendamos usar uma aba chamada 'Dados'")
                
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                df.columns = df.columns.str.strip().str.upper()
                
                st.success(f"Dados carregados: {len(df)} linhas, {len(df.columns)} colunas")
                
                with st.expander("Preview dos Dados", expanded=True):
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Linhas", len(df))
                    with col2:
                        st.metric("Colunas", len(df.columns))
                    with col3:
                        if "LOJA" in df.columns:
                            st.metric("Lojas", df["LOJA"].dropna().nunique())
                
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {str(e)}")
            logger.error(f"Erro leitura: {e}")
            st.stop()

    if df is not None:
        st.markdown("### Validação dos Dados")
        
        with st.spinner("Validando dados..."):
            erros, avisos = validar_dados_enviados(df)
        
        if erros:
            st.error("Problemas encontrados:")
            for erro in erros:
                st.error(f"- {erro}")
            st.stop()
        else:
            st.success("Validação aprovada!")
        
        if avisos:
            for aviso in avisos:
                st.info(aviso)
        
        st.divider()
        
        # Botões
        col1, col2 = st.columns([2, 1])
        
        with col1:
            if st.button("Consolidar Dados", type="primary", use_container_width=True):
                st.warning("Consolidação iniciada! NÃO feche esta página!")
                
                sucesso = processar_consolidacao_com_lock(df, uploaded_file.name, token)
                
                if sucesso:
                    st.balloons()
        
        with col2:
            if st.button("Limpar Tela", type="secondary", use_container_width=True):
                st.rerun()

    # Footer
    st.markdown("---")
    st.markdown(f"""
    <div style="text-align: center; padding: 1rem; color: #666;">
        <strong>Sistema de Consolidação de Bonificações v{APP_VERSION}</strong><br>
        <small>Última atualização: {VERSION_DATE}</small>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
