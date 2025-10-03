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
APP_VERSION = "1.1.0"
VERSION_DATE = "2025-10-03"
APP_TITLE = "Sistema de Bonificações"
APP_SUBTITLE = "Substituição completa do arquivo consolidado"

# ===========================
# CONFIGURAÇÃO DE LOGGING
# ===========================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ===========================
# CREDENCIAIS VIA ST.SECRETS
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
    logger.info("Credenciais carregadas com sucesso")
except KeyError as e:
    CREDENCIAL_FALTANDO = str(e)
    logger.error(f"Credencial faltando: {e}")

# ===========================
# CONFIGURAÇÃO DE PASTAS
# ===========================
PASTA_CONSOLIDADO = "Documentos Compartilhados/LimparAuto/FontedeDados"
PASTA_ENVIOS_BACKUPS = "Documentos Compartilhados/PlanilhasEnviadas_Backups/Bonificacao"
ARQUIVO_LOCK = "sistema_lock_bonificacao.json"
TIMEOUT_LOCK_MINUTOS = 10

# ===========================
# ESTILOS CSS
# ===========================
def aplicar_estilos_css():
    """Aplica estilos CSS customizados"""
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
    
    .metric-box {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    </style>
    """, unsafe_allow_html=True)

# ===========================
# AUTENTICAÇÃO
# ===========================
def obter_token():
    """
    Obtém token de acesso do Microsoft Graph API
    Usa session_state para cache ao invés de st.cache_data
    """
    if 'token_cache' not in st.session_state or \
       'token_timestamp' not in st.session_state or \
       (datetime.now() - st.session_state.token_timestamp).seconds > 3000:
        
        try:
            app = ConfidentialClientApplication(
                CLIENT_ID,
                authority=f"https://login.microsoftonline.com/{TENANT_ID}",
                client_credential=CLIENT_SECRET
            )
            result = app.acquire_token_for_client(
                scopes=["https://graph.microsoft.com/.default"]
            )
            
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
    """Gera ID único para a sessão"""
    if 'session_id' not in st.session_state:
        st.session_state.session_id = str(uuid.uuid4())[:8]
    return st.session_state.session_id

def verificar_lock_existente(token):
    """Verifica se existe um lock ativo no sistema"""
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{ARQUIVO_LOCK}:/content"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            lock_data = response.json()
            timestamp_lock = datetime.fromisoformat(lock_data['timestamp'])
            agora = datetime.now()
            
            if agora - timestamp_lock > timedelta(minutes=TIMEOUT_LOCK_MINUTOS):
                logger.info("Lock expirado - removendo automaticamente")
                remover_lock(token, force=True)
                return False, None
            
            return True, lock_data
        
        return False, None
            
    except Exception as e:
        logger.error(f"Erro ao verificar lock: {e}")
        return False, None

def criar_lock(token, operacao="Substituição completa"):
    """Cria um lock para bloquear outras operações"""
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
        
        response = requests.put(
            url, 
            headers=headers, 
            data=json.dumps(lock_data), 
            timeout=10
        )
        
        if response.status_code in [200, 201]:
            logger.info(f"Lock criado com sucesso. Session ID: {session_id}")
            return True, session_id
        
        logger.error(f"Erro ao criar lock: {response.status_code}")
        return False, None
            
    except Exception as e:
        logger.error(f"Erro ao criar lock: {e}")
        return False, None

def remover_lock(token, session_id=None, force=False):
    """Remove o lock do sistema"""
    try:
        if not force and session_id:
            lock_existe, lock_data = verificar_lock_existente(token)
            if lock_existe and lock_data.get('session_id') != session_id:
                logger.warning("Tentativa de remover lock de outra sessão")
                return False
        
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{ARQUIVO_LOCK}"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.delete(url, headers=headers, timeout=10)
        
        if response.status_code in [200, 204, 404]:
            logger.info("Lock removido com sucesso")
            return True
        
        logger.error(f"Erro ao remover lock: {response.status_code}")
        return False
            
    except Exception as e:
        logger.error(f"Erro ao remover lock: {e}")
        return False

def exibir_status_sistema(token):
    """Exibe o status atual do sistema de lock"""
    lock_existe, lock_data = verificar_lock_existente(token)
    
    if lock_existe:
        timestamp_inicio = datetime.fromisoformat(lock_data['timestamp'])
        duracao = datetime.now() - timestamp_inicio
        
        st.markdown("""
        <div class="status-card error">
            <h3>Sistema Ocupado</h3>
            <p>Outro usuário está processando dados no momento</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Tempo Ativo", f"{int(duracao.total_seconds()//60)} min")
        with col2:
            st.metric("Operação", lock_data.get('operacao', 'N/A'))
        
        tempo_limite = timestamp_inicio + timedelta(minutes=TIMEOUT_LOCK_MINUTOS)
        tempo_restante = tempo_limite - datetime.now()
        
        if tempo_restante.total_seconds() < 0:
            if st.button("Liberar Sistema (Forçar)", type="secondary"):
                if remover_lock(token, force=True):
                    st.success("Sistema liberado com sucesso")
                    st.rerun()
        
        return True
    else:
        st.markdown("""
        <div class="status-card success">
            <h3>Sistema Disponível</h3>
            <p>Pronto para receber sua planilha</p>
        </div>
        """, unsafe_allow_html=True)
        return False

# ===========================
# FUNÇÕES AUXILIARES
# ===========================
def criar_pasta_se_nao_existir(caminho_pasta, token):
    """Cria estrutura de pastas no OneDrive se não existir"""
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
                
                create_response = requests.post(
                    parent_url, 
                    headers={**headers, "Content-Type": "application/json"}, 
                    json=create_body,
                    timeout=10
                )
                
                if create_response.status_code in [200, 201]:
                    logger.info(f"Pasta criada: {parte}")
                    
    except Exception as e:
        logger.warning(f"Erro ao criar pastas: {e}")

def upload_onedrive(nome_arquivo, conteudo_arquivo, token, tipo_arquivo="consolidado"):
    """Faz upload de arquivo para OneDrive"""
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
        
        if response.status_code in [200, 201]:
            logger.info(f"Upload realizado com sucesso: {nome_arquivo}")
            return True
        else:
            logger.error(f"Erro no upload: {response.status_code}")
            return False
        
    except Exception as e:
        logger.error(f"Erro no upload: {e}")
        return False

def baixar_arquivo_consolidado(token):
    """Baixa o arquivo consolidado existente"""
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
        logger.error(f"Erro ao baixar arquivo consolidado: {e}")
        return pd.DataFrame(), False

def salvar_arquivo_enviado(df_novo, nome_arquivo_original, token):
    """Salva uma cópia do arquivo enviado na pasta de backups"""
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d_%Hh%M")
        nome_base = nome_arquivo_original.replace(".xlsx", "").replace(".xls", "")
        nome_arquivo_backup = f"{nome_base}_enviado_{timestamp}.xlsx"
        
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_novo.to_excel(writer, index=False, sheet_name="Dados")
        buffer.seek(0)
        
        sucesso = upload_onedrive(nome_arquivo_backup, buffer.read(), token, "backup")
        
        if sucesso:
            logger.info(f"Backup do arquivo enviado salvo: {nome_arquivo_backup}")
        else:
            logger.warning(f"Não foi possível salvar backup do arquivo enviado")
            
    except Exception as e:
        logger.error(f"Erro ao salvar arquivo enviado: {e}")

def fazer_backup_consolidado(token):
    """Faz backup do arquivo consolidado atual antes de substituir"""
    try:
        df_antigo, existe = baixar_arquivo_consolidado(token)
        
        if existe and not df_antigo.empty:
            timestamp = datetime.now().strftime("%Y-%m-%d_%Hh%M")
            nome_arquivo_backup = f"bonificacao_consolidada_backup_{timestamp}.xlsx"
            
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_antigo.to_excel(writer, index=False, sheet_name="Dados")
            buffer.seek(0)
            
            sucesso = upload_onedrive(nome_arquivo_backup, buffer.read(), token, "backup")
            
            if sucesso:
                logger.info(f"Backup do consolidado criado: {nome_arquivo_backup}")
                return True, len(df_antigo)
            else:
                logger.warning("Não foi possível criar backup do consolidado")
                return False, 0
        
        logger.info("Nenhum arquivo consolidado anterior para fazer backup")
        return False, 0
            
    except Exception as e:
        logger.error(f"Erro ao fazer backup do consolidado: {e}")
        return False, 0

# ===========================
# VALIDAÇÃO
# ===========================
def validar_dados_enviados(df):
    """Validação rigorosa dos dados enviados"""
    erros = []
    avisos = []
    
    if df.empty:
        erros.append("A planilha está vazia")
        return erros, avisos
    
    # Validar coluna LOJA
    if "LOJA" not in df.columns:
        erros.append("A planilha deve conter uma coluna 'LOJA'")
    else:
        lojas_validas = df["LOJA"].notna().sum()
        if lojas_validas == 0:
            erros.append("Nenhuma loja válida encontrada na coluna 'LOJA'")
        else:
            lojas_unicas = df["LOJA"].dropna().unique()
            avisos.append(f"Lojas encontradas: {len(lojas_unicas)}")
    
    # Validar coluna DATA
    if "DATA" not in df.columns:
        erros.append("A planilha deve conter uma coluna 'DATA'")
    else:
        try:
            df_temp = df.copy()
            df_temp["DATA"] = pd.to_datetime(df_temp["DATA"], errors="coerce")
            datas_invalidas = df_temp["DATA"].isna().sum()
            
            if datas_invalidas > 0:
                avisos.append(f"{datas_invalidas} datas inválidas serão removidas")
            
            if datas_invalidas == len(df_temp):
                erros.append("Todas as datas são inválidas")
        except Exception as e:
            erros.append(f"Erro ao processar datas: {str(e)}")
    
    return erros, avisos

# ===========================
# CONSOLIDAÇÃO - CENÁRIO B
# ===========================
def processar_substituicao_completa(df_novo, nome_arquivo, token):
    """
    CENÁRIO B: Substitui COMPLETAMENTE o arquivo consolidado
    - Deleta todos os dados anteriores
    - Cria arquivo novo com apenas os dados enviados
    - Adiciona campo DATA_ULTIMO_ENVIO
    - Faz backup automático antes de substituir
    """
    session_id = gerar_id_sessao()
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        status_text.info("Iniciando substituição completa...")
        progress_bar.progress(5)
        
        # Verificar lock
        sistema_ocupado, _ = verificar_lock_existente(token)
        if sistema_ocupado:
            status_text.error("Sistema ocupado por outro usuário")
            return False
        
        # Criar lock
        status_text.info("Bloqueando sistema...")
        progress_bar.progress(10)
        
        lock_criado, session_lock = criar_lock(token, "Substituição completa do consolidado")
        if not lock_criado:
            status_text.error("Não foi possível bloquear o sistema")
            return False
        
        # Fazer backup do consolidado atual
        status_text.info("Fazendo backup do arquivo atual...")
        progress_bar.progress(20)
        
        backup_feito, registros_antigos = fazer_backup_consolidado(token)
        if backup_feito:
            st.info(f"Backup criado: {registros_antigos} registros salvos")
        
        # Preparar dados novos
        status_text.info("Preparando dados...")
        progress_bar.progress(35)
        
        df_novo = df_novo.copy()
        df_novo.columns = df_novo.columns.str.strip().str.upper()
        
        # Converter datas
        df_novo["DATA"] = pd.to_datetime(df_novo["DATA"], errors="coerce")
        linhas_invalidas = df_novo["DATA"].isna().sum()
        df_novo = df_novo.dropna(subset=["DATA"])
        
        if df_novo.empty:
            status_text.error("Nenhum registro válido após limpeza de datas")
            remover_lock(token, session_lock)
            return False
        
        if linhas_invalidas > 0:
            st.warning(f"{linhas_invalidas} linhas com datas inválidas foram removidas")
        
        # ADICIONAR CAMPO DATA_ULTIMO_ENVIO
        status_text.info("Adicionando campo DATA_ULTIMO_ENVIO...")
        progress_bar.progress(50)
        
        data_envio = datetime.now()
        df_novo['DATA_ULTIMO_ENVIO'] = data_envio
        
        logger.info(f"Campo DATA_ULTIMO_ENVIO adicionado: {data_envio.strftime('%d/%m/%Y %H:%M:%S')}")
        
        # Ordenar dados
        df_novo = df_novo.sort_values(
            ["DATA", "LOJA"], 
            na_position='last'
        ).reset_index(drop=True)
        
        # Salvar cópia do arquivo enviado
        status_text.info("Salvando cópia do arquivo enviado...")
        progress_bar.progress(65)
        salvar_arquivo_enviado(df_novo, nome_arquivo, token)
        
        # Upload do novo consolidado (SUBSTITUIÇÃO TOTAL)
        status_text.info("Salvando novo arquivo consolidado...")
        progress_bar.progress(80)
        
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_novo.to_excel(writer, index=False, sheet_name="Dados")
        buffer.seek(0)
        
        sucesso = upload_onedrive("bonificacao_consolidada.xlsx", buffer.read(), token, "consolidado")
        
        # Remover lock
        progress_bar.progress(95)
        remover_lock(token, session_lock)
        progress_bar.progress(100)
        
        if sucesso:
            status_text.empty()
            progress_bar.empty()
            
            # Mensagem de sucesso
            st.success("Substituição completa realizada com sucesso!")
            
            if backup_feito:
                st.warning(f"O arquivo consolidado anterior ({registros_antigos} registros) foi substituído e salvo em backup")
            else:
                st.info("Arquivo consolidado criado (não havia arquivo anterior)")
            
            # Métricas do resultado
            st.markdown("### Resultado da Operação")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total de Registros", f"{len(df_novo):,}")
            
            with col2:
                if "LOJA" in df_novo.columns:
                    lojas_total = df_novo["LOJA"].dropna().nunique()
                    st.metric("Total de Lojas", lojas_total)
            
            with col3:
                st.metric("Data do Envio", data_envio.strftime("%d/%m/%Y %H:%M"))
            
            # Informação sobre o campo DATA_ULTIMO_ENVIO
            st.info(f"Campo 'DATA_ULTIMO_ENVIO' adicionado com sucesso em todos os {len(df_novo)} registros")
            
            # Resumo por loja
            if not df_novo.empty and 'LOJA' in df_novo.columns:
                st.markdown("### Resumo por Loja")
                
                resumo = df_novo.groupby("LOJA").agg({
                    "DATA": ["count", "min", "max"]
                })
                resumo.columns = ["Total Registros", "Data Inicial", "Data Final"]
                
                # Formatar datas
                resumo["Data Inicial"] = pd.to_datetime(resumo["Data Inicial"]).dt.strftime("%d/%m/%Y")
                resumo["Data Final"] = pd.to_datetime(resumo["Data Final"]).dt.strftime("%d/%m/%Y")
                
                st.dataframe(resumo, use_container_width=True)
            
            # Localização dos arquivos
            with st.expander("Localização dos Arquivos"):
                st.info(f"**Arquivo Consolidado:**\n`{PASTA_CONSOLIDADO}/bonificacao_consolidada.xlsx`")
                st.info(f"**Backups:**\n`{PASTA_ENVIOS_BACKUPS}/`")
            
            return True
        else:
            status_text.error("Erro ao salvar arquivo consolidado")
            return False
            
    except Exception as e:
        logger.error(f"Erro na substituição completa: {e}")
        remover_lock(token, session_id, force=True)
        status_text.error(f"Erro durante o processo: {str(e)}")
        progress_bar.empty()
        st.error("Sistema liberado automaticamente após erro")
        return False

# ===========================
# INTERFACE PRINCIPAL
# ===========================
def main():
    """Função principal da aplicação"""
    
    st.set_page_config(
        page_title=f"{APP_TITLE} v{APP_VERSION}", 
        layout="wide",
        initial_sidebar_state="expanded",
        page_icon="📊"
    )

    aplicar_estilos_css()

    # Header principal
    st.markdown(f"""
    <div class="main-header">
        <h1>{APP_TITLE}</h1>
        <p style="margin: 0.5rem 0 0 0; opacity: 0.9;">{APP_SUBTITLE}</p>
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
        st.info("Vá em: Manage app → Settings → Secrets")
        st.stop()

    # Sidebar
    st.sidebar.markdown("### Upload de Bonificações")
    st.sidebar.markdown(f"**Versão:** {APP_VERSION}")
    st.sidebar.divider()

    # Obter token
    with st.spinner("Conectando ao Microsoft Graph..."):
        token = obter_token()
    
    if not token:
        st.error("Erro de autenticação. Verifique as credenciais nas secrets.")
        st.sidebar.error("❌ Desconectado")
        st.stop()
    
    st.sidebar.success("✅ Conectado")

    # Status do sistema
    st.markdown("## Status do Sistema")
    sistema_ocupado = exibir_status_sistema(token)
    
    if sistema_ocupado:
        st.divider()
        if st.button("🔄 Atualizar Status"):
            st.rerun()
        st.info("Página será atualizada automaticamente em 15 segundos")
        time.sleep(15)
        st.rerun()

    st.divider()

    # Avisos importantes
    st.warning("⚠️ **ATENÇÃO:** Este sistema faz SUBSTITUIÇÃO COMPLETA do arquivo consolidado. Todos os dados anteriores serão substituídos pelos novos dados enviados.")
    
    st.info("""**Funcionalidades:**
- ✅ Substitui completamente o arquivo consolidado
- ✅ Faz backup automático antes de substituir
- ✅ Adiciona campo DATA_ULTIMO_ENVIO em todos os registros
- ✅ Salva cópia do arquivo enviado
    """)

    # Informações do sistema
    with st.sidebar.expander("ℹ️ Informações"):
        st.markdown(f"**Modo:** Substituição Completa")
        st.markdown(f"**Consolidado:** bonificacao_consolidada.xlsx")
        st.markdown(f"**Pasta:** {PASTA_CONSOLIDADO}")

    # Upload de arquivo
    st.markdown("## Upload de Planilha Excel")
    
    st.info("📋 A planilha deve ter uma aba 'Dados' com as colunas 'LOJA' e 'DATA'")

    uploaded_file = st.file_uploader(
        "Escolha um arquivo Excel", 
        type=["xlsx", "xls"],
        help="Formatos aceitos: .xlsx, .xls"
    )

    df = None
    if uploaded_file:
        try:
            st.success(f"📁 Arquivo carregado: {uploaded_file.name}")
            
            with st.spinner("📖 Lendo arquivo..."):
                xls = pd.ExcelFile(uploaded_file)
                sheets = xls.sheet_names
                
                if "Dados" in sheets:
                    sheet = "Dados"
                    st.success("✅ Aba 'Dados' encontrada automaticamente")
                else:
                    sheet = st.selectbox("Selecione a aba:", sheets)
                    if sheet != "Dados":
                        st.warning("⚠️ Recomendamos usar uma aba chamada 'Dados'")
                
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                df.columns = df.columns.str.strip().str.upper()
                
                st.success(f"✅ Dados carregados: {len(df)} linhas, {len(df.columns)} colunas")
                
                # Preview dos dados
                with st.expander("👀 Preview dos Dados", expanded=True):
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Linhas", len(df))
                    with col2:
                        st.metric("Colunas", len(df.columns))
                    with col3:
                        if "LOJA" in df.columns:
                            st.metric("Lojas", df["LOJA"].dropna().nunique())
                        else:
                            st.metric("Lojas", "N/A")
                
        except Exception as e:
            st.error(f"❌ Erro ao ler arquivo: {str(e)}")
            logger.error(f"Erro na leitura do arquivo: {e}")
            st.stop()

    # Validação e processamento
    if df is not None:
        st.markdown("### 🔍 Validação dos Dados")
        
        with st.spinner("Validando dados..."):
            erros, avisos = validar_dados_enviados(df)
        
        if erros:
            st.error("❌ Problemas encontrados:")
            for erro in erros:
                st.error(f"• {erro}")
            st.stop()
        else:
            st.success("✅ Validação aprovada!")
        
        if avisos:
            for aviso in avisos:
                st.info(f"ℹ️ {aviso}")
        
        st.divider()
        
        # Botões de ação
        col1, col2 = st.columns([2, 1])
        
        with col1:
            if st.button("⚠️ Substituir Arquivo Consolidado", type="primary", use_container_width=True):
                st.warning("⏳ Substituição iniciada! NÃO feche esta página!")
                
                sucesso = processar_substituicao_completa(df, uploaded_file.name, token)
                
                if sucesso:
                    st.balloons()
                    st.success("🎉 Processo concluído com sucesso!")
        
        with col2:
            if st.button("🔄 Limpar Tela", type="secondary", use_container_width=True):
                st.rerun()

    # Footer
    st.markdown("---")
    st.markdown(f"""
    <div style="text-align: center; padding: 1rem; color: #666;">
        <strong>{APP_TITLE} v{APP_VERSION}</strong><br>
        <small>Modo: Substituição Completa | Última atualização: {VERSION_DATE}</small>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
