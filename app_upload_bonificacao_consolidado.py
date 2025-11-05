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
APP_VERSION = "2.0.0"
VERSION_DATE = "2025-11-05"
APP_TITLE = "Upload da planilha de  Bonificações"
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
# ESTRUTURA DE COLUNAS ESPERADA
# ===========================
COLUNAS_OBRIGATORIAS = [
    'GRUPO', 'CONCESSIONÁRIA', 'LOJA', 'FUNÇÃO', 'NOME', 'DATA', 
    'FORMA PAG', 'CARTÃO / PIX',
    'TMO_DUTO', 'R$_DUTO', 'EXTRA_DUTO', 'TOTAL_DUTO',
    'TMO_FREIO', 'R$_FREIO', 'EXTRA_FREIO', 'TOTAL_FREIO',
    'TMO_SANIT', 'R$_SANIT', 'EXTRA_SANIT', 'TOTAL_SANIT',
    'TMO_VERNIZ', 'R$_VERNIZ', 'EXTRA_VERNIZ', 'TOTAL_VERNIZ',
    'TMO_CX EVAP', 'R$_CX EVAP', 'EXTRA_CX EVAP', 'TOTAL_CX EVAP',
    'TMO_PROTEC', 'R$_PROTEC', 'EXTRA_PROTEC', 'TOTAL_PROTEC',
    'TMO_VC GREEN', 'R$_VC GREEN', 'EXTRA_VC GREEN', 'TOTAL_VC GREEN',
    'TMO_NITROGÊNIO', 'R$_NITROGÊNIO', 'EXTRA_NITROGÊNIO', 'TOTAL_NITROGÊNIO',
    'TMO_TOTAL', 'R$_TOTAL', 'STATUS', 'PAGO', 'A PAGAR', 'PIX'
]

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
    
    .validation-box {
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    
    .column-list {
        max-height: 300px;
        overflow-y: auto;
        background: white;
        border: 1px solid #e0e0e0;
        border-radius: 5px;
        padding: 0.5rem;
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
    """Exibe status atual do sistema"""
    lock_existe, lock_data = verificar_lock_existente(token)
    
    if lock_existe:
        st.markdown('<div class="status-card error">', unsafe_allow_html=True)
        st.error("🔒 Sistema ocupado - Operação em andamento")
        
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"**Operação:** {lock_data.get('operacao', 'Desconhecida')}")
            st.info(f"**Sessão ID:** {lock_data.get('session_id', 'N/A')}")
        
        with col2:
            timestamp = datetime.fromisoformat(lock_data.get('timestamp'))
            tempo_decorrido = datetime.now() - timestamp
            minutos = int(tempo_decorrido.seconds / 60)
            st.info(f"**Iniciado há:** {minutos} minuto(s)")
            st.info(f"**Versão:** {lock_data.get('app_version', 'N/A')}")
        
        st.markdown('</div>', unsafe_allow_html=True)
        return True
    else:
        st.markdown('<div class="status-card success">', unsafe_allow_html=True)
        st.success("✅ Sistema disponível")
        st.markdown('</div>', unsafe_allow_html=True)
        return False

# ===========================
# VALIDAÇÃO DE ESTRUTURA
# ===========================
def baixar_arquivo_consolidado(token):
    """Baixa o arquivo consolidado atual para comparação de estrutura"""
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/bonificacao_consolidada.xlsx:/content"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            return pd.read_excel(BytesIO(response.content), sheet_name='Dados')
        else:
            logger.warning(f"Arquivo consolidado não encontrado: {response.status_code}")
            return None
            
    except Exception as e:
        logger.error(f"Erro ao baixar arquivo consolidado: {e}")
        return None

def validar_estrutura_colunas(df_novo, token):
    """
    Valida a estrutura de colunas do arquivo enviado
    Retorna: (erros, avisos, info_validacao)
    """
    erros = []
    avisos = []
    info_validacao = {
        'colunas_faltando': [],
        'colunas_diferentes': [],
        'colunas_novas': [],
        'estrutura_ok': False
    }
    
    # Normalizar colunas do novo arquivo
    colunas_novo = [col.strip().upper() for col in df_novo.columns]
    
    # 1. Verificar colunas obrigatórias
    colunas_faltando = [col for col in COLUNAS_OBRIGATORIAS if col not in colunas_novo]
    
    if colunas_faltando:
        erros.append(f"❌ Colunas obrigatórias faltando: {', '.join(colunas_faltando)}")
        info_validacao['colunas_faltando'] = colunas_faltando
        return erros, avisos, info_validacao
    
    # 2. Baixar arquivo consolidado atual para comparar estrutura
    df_consolidado = baixar_arquivo_consolidado(token)
    
    if df_consolidado is not None:
        colunas_consolidado = [col.strip().upper() for col in df_consolidado.columns]
        
        # Remover DATA_ULTIMO_ENVIO da comparação (é adicionada automaticamente)
        colunas_consolidado_sem_data = [col for col in colunas_consolidado if col != 'DATA_ULTIMO_ENVIO']
        
        # 3. Verificar mudanças nos nomes de colunas existentes
        colunas_diferentes = []
        for col in colunas_consolidado_sem_data:
            if col in COLUNAS_OBRIGATORIAS and col not in colunas_novo:
                colunas_diferentes.append(col)
        
        if colunas_diferentes:
            erros.append(f"❌ As seguintes colunas mudaram de nome ou estão ausentes: {', '.join(colunas_diferentes)}")
            info_validacao['colunas_diferentes'] = colunas_diferentes
            return erros, avisos, info_validacao
        
        # 4. Identificar novas colunas (permitidas)
        colunas_novas = [col for col in colunas_novo if col not in colunas_consolidado_sem_data and col != 'DATA_ULTIMO_ENVIO']
        
        if colunas_novas:
            avisos.append(f"ℹ️ Novas colunas detectadas (serão adicionadas): {', '.join(colunas_novas)}")
            info_validacao['colunas_novas'] = colunas_novas
    
    else:
        # Arquivo consolidado não existe ainda, apenas validar colunas obrigatórias
        avisos.append("⚠️ Arquivo consolidado não existe. Será criado pela primeira vez.")
        
        # Verificar se há colunas além das obrigatórias
        colunas_extras = [col for col in colunas_novo if col not in COLUNAS_OBRIGATORIAS and col != 'DATA_ULTIMO_ENVIO']
        if colunas_extras:
            avisos.append(f"ℹ️ Colunas adicionais no arquivo: {', '.join(colunas_extras)}")
            info_validacao['colunas_novas'] = colunas_extras
    
    info_validacao['estrutura_ok'] = True
    return erros, avisos, info_validacao

def validar_dados_enviados(df, token):
    """
    Valida os dados enviados
    Retorna: (erros, avisos)
    """
    erros = []
    avisos = []
    
    # Normalizar nomes das colunas
    df.columns = df.columns.str.strip().str.upper()
    
    # 1. Validar estrutura de colunas
    erros_estrutura, avisos_estrutura, info_validacao = validar_estrutura_colunas(df, token)
    erros.extend(erros_estrutura)
    avisos.extend(avisos_estrutura)
    
    if not info_validacao['estrutura_ok']:
        return erros, avisos
    
    # 2. Verificar se DataFrame está vazio
    if df.empty:
        erros.append("Planilha está vazia")
        return erros, avisos
    
    # 3. Verificar colunas essenciais para operação
    if "LOJA" not in df.columns:
        erros.append("Coluna 'LOJA' não encontrada")
    
    if "DATA" not in df.columns:
        erros.append("Coluna 'DATA' não encontrada")
    
    # 4. Validações adicionais se as colunas existem
    if "LOJA" in df.columns:
        lojas_vazias = df["LOJA"].isna().sum()
        if lojas_vazias > 0:
            avisos.append(f"Existem {lojas_vazias} linha(s) sem LOJA identificada")
    
    if "DATA" in df.columns:
        datas_vazias = df["DATA"].isna().sum()
        if datas_vazias > 0:
            avisos.append(f"Existem {datas_vazias} linha(s) sem DATA")
        
        # Tentar converter DATA para datetime
        try:
            df["DATA"] = pd.to_datetime(df["DATA"], errors='coerce')
            datas_invalidas = df["DATA"].isna().sum()
            if datas_invalidas > 0:
                avisos.append(f"Existem {datas_invalidas} data(s) inválida(s)")
        except Exception as e:
            erros.append(f"Erro ao processar datas: {str(e)}")
    
    return erros, avisos

# ===========================
# FUNÇÕES DE ARQUIVO
# ===========================
def salvar_arquivo_sharepoint(df, nome_arquivo, pasta, token):
    """Salva DataFrame como Excel no SharePoint"""
    try:
        buffer = BytesIO()
        
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Dados', index=False)
        
        buffer.seek(0)
        conteudo = buffer.read()
        
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{pasta}/{nome_arquivo}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        
        response = requests.put(url, headers=headers, data=conteudo, timeout=60)
        
        if response.status_code in [200, 201]:
            logger.info(f"Arquivo salvo: {nome_arquivo}")
            return True
        else:
            logger.error(f"Erro ao salvar {nome_arquivo}: {response.status_code}")
            return False
            
    except Exception as e:
        logger.error(f"Erro ao salvar arquivo: {e}")
        return False

# ===========================
# PROCESSAMENTO
# ===========================
def processar_substituicao_completa(df, nome_arquivo_original, token):
    """
    Processa substituição completa do arquivo consolidado
    """
    try:
        # 1. Criar lock
        sucesso_lock, session_id = criar_lock(token, "Substituição completa")
        if not sucesso_lock:
            st.error("❌ Não foi possível criar lock. Sistema pode estar ocupado.")
            return False
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 2. Backup do arquivo atual
        status_text.info("📦 Fazendo backup do arquivo atual...")
        progress_bar.progress(20)
        
        df_atual = baixar_arquivo_consolidado(token)
        if df_atual is not None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_backup = f"backup_bonificacao_{timestamp}.xlsx"
            
            if not salvar_arquivo_sharepoint(df_atual, nome_backup, PASTA_ENVIOS_BACKUPS, token):
                st.warning("⚠️ Backup não foi salvo, mas continuando...")
            else:
                st.success(f"✅ Backup salvo: {nome_backup}")
        else:
            st.info("ℹ️ Nenhum arquivo anterior para backup")
        
        progress_bar.progress(40)
        
        # 3. Preparar novos dados
        status_text.info("🔄 Preparando novos dados...")
        df_novo = df.copy()
        
        # Garantir que colunas estão em maiúsculas
        df_novo.columns = df_novo.columns.str.strip().str.upper()
        
        # Adicionar/atualizar DATA_ULTIMO_ENVIO
        data_envio = datetime.now()
        df_novo["DATA_ULTIMO_ENVIO"] = data_envio
        
        progress_bar.progress(60)
        
        # 4. Salvar cópia do arquivo enviado
        status_text.info("💾 Salvando cópia do arquivo enviado...")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_copia = f"enviado_{timestamp}_{nome_arquivo_original}"
        
        salvar_arquivo_sharepoint(df_novo, nome_copia, PASTA_ENVIOS_BACKUPS, token)
        st.success(f"✅ Cópia salva: {nome_copia}")
        
        progress_bar.progress(80)
        
        # 5. Salvar arquivo consolidado
        status_text.info("💾 Salvando arquivo consolidado...")
        
        sucesso = salvar_arquivo_sharepoint(
            df_novo,
            "bonificacao_consolidada.xlsx",
            PASTA_CONSOLIDADO,
            token
        )
        
        progress_bar.progress(100)
        
        # 6. Remover lock
        remover_lock(token, session_id)
        
        if sucesso:
            status_text.success("✅ Arquivo consolidado atualizado com sucesso!")
            progress_bar.empty()
            
            # Exibir resumo
            st.markdown("### 📊 Resumo da Substituição")
            
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
            st.info(f"✅ Campo 'DATA_ULTIMO_ENVIO' adicionado em todos os {len(df_novo)} registros")
            
            # Exibir informações sobre colunas se houver novas
            if 'info_validacao' in st.session_state and st.session_state.info_validacao.get('colunas_novas'):
                st.success(f"✅ Novas colunas adicionadas: {', '.join(st.session_state.info_validacao['colunas_novas'])}")
            
            # Resumo por loja
            if not df_novo.empty and 'LOJA' in df_novo.columns:
                st.markdown("### 📋 Resumo por Loja")
                
                resumo = df_novo.groupby("LOJA").agg({
                    "DATA": ["count", "min", "max"]
                })
                resumo.columns = ["Total Registros", "Data Inicial", "Data Final"]
                
                # Formatar datas
                resumo["Data Inicial"] = pd.to_datetime(resumo["Data Inicial"]).dt.strftime("%d/%m/%Y")
                resumo["Data Final"] = pd.to_datetime(resumo["Data Final"]).dt.strftime("%d/%m/%Y")
                
                st.dataframe(resumo, use_container_width=True)
            
            # Localização dos arquivos
            with st.expander("📁 Localização dos Arquivos"):
                st.info(f"**Arquivo Consolidado:**\n`{PASTA_CONSOLIDADO}/bonificacao_consolidada.xlsx`")
                st.info(f"**Backups:**\n`{PASTA_ENVIOS_BACKUPS}/`")
            
            return True
        else:
            status_text.error("❌ Erro ao salvar arquivo consolidado")
            return False
            
    except Exception as e:
        logger.error(f"Erro na substituição completa: {e}")
        remover_lock(token, session_id, force=True)
        status_text.error(f"❌ Erro durante o processo: {str(e)}")
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
    st.sidebar.markdown("### 📤 Upload de Bonificações")
    st.sidebar.markdown(f"**Versão:** {APP_VERSION}")
    st.sidebar.divider()

    # Obter token
    with st.spinner("🔄 Conectando ao Microsoft Graph..."):
        token = obter_token()
    
    if not token:
        st.error("❌ Erro de autenticação. Verifique as credenciais nas secrets.")
        st.sidebar.error("❌ Desconectado")
        st.stop()
    
    st.sidebar.success("✅ Conectado")

    # Status do sistema
    st.markdown("## 🔍 Status do Sistema")
    sistema_ocupado = exibir_status_sistema(token)
    
    if sistema_ocupado:
        st.divider()
        if st.button("🔄 Atualizar Status"):
            st.rerun()
        st.info("⏱️ Página será atualizada automaticamente em 15 segundos")
        time.sleep(15)
        st.rerun()

    st.divider()

    # Avisos importantes
    st.warning("⚠️ **ATENÇÃO:** Este sistema faz SUBSTITUIÇÃO COMPLETA do arquivo consolidado. Todos os dados anteriores serão substituídos pelos novos dados enviados.")
    
    st.info("""**✨ Funcionalidades:**
- ✅ Substitui completamente o arquivo consolidado
- ✅ Valida estrutura de colunas obrigatórias
- ✅ Permite adição de novas colunas
- ✅ Detecta mudanças nos nomes de colunas
- ✅ Faz backup automático antes de substituir
- ✅ Adiciona campo DATA_ULTIMO_ENVIO em todos os registros
- ✅ Salva cópia do arquivo enviado
    """)

    # Informações do sistema
    with st.sidebar.expander("ℹ️ Informações"):
        st.markdown(f"**Modo:** Substituição Completa")
        st.markdown(f"**Consolidado:** bonificacao_consolidada.xlsx")
        st.markdown(f"**Pasta:** {PASTA_CONSOLIDADO}")
        
        with st.expander("📋 Colunas Obrigatórias"):
            st.markdown('<div class="column-list">', unsafe_allow_html=True)
            for col in COLUNAS_OBRIGATORIAS:
                st.text(f"• {col}")
            st.markdown('</div>', unsafe_allow_html=True)

    # Upload de arquivo
    st.markdown("## 📤 Upload de Planilha Excel")
    
    st.info("📋 A planilha deve ter uma aba 'Dados' com todas as colunas obrigatórias")

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
        
        with st.spinner("🔄 Validando dados e estrutura..."):
            erros, avisos = validar_dados_enviados(df, token)
        
        if erros:
            st.error("❌ **Problemas encontrados:**")
            for erro in erros:
                st.error(f"• {erro}")
            
            # Mostrar comparação de colunas se houver erro de estrutura
            st.markdown("---")
            st.markdown("### 📋 Comparação de Estrutura")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Colunas no seu arquivo:**")
                st.markdown('<div class="validation-box">', unsafe_allow_html=True)
                colunas_usuario = [col for col in df.columns if col != 'DATA_ULTIMO_ENVIO']
                for col in sorted(colunas_usuario):
                    st.text(f"• {col}")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col2:
                st.markdown("**Colunas obrigatórias:**")
                st.markdown('<div class="validation-box">', unsafe_allow_html=True)
                for col in sorted(COLUNAS_OBRIGATORIAS):
                    if col in df.columns:
                        st.text(f"✅ {col}")
                    else:
                        st.text(f"❌ {col}")
                st.markdown('</div>', unsafe_allow_html=True)
            
            st.stop()
        else:
            st.success("✅ **Validação aprovada!**")
        
        if avisos:
            st.markdown("### ℹ️ Informações Adicionais")
            for aviso in avisos:
                st.info(aviso)
        
        # Guardar info de validação para usar depois
        erros_estrutura, avisos_estrutura, info_validacao = validar_estrutura_colunas(df, token)
        st.session_state.info_validacao = info_validacao
        
        st.divider()
        
        # Botões de ação
        col1, col2 = st.columns([2, 1])
        
        with col1:
            if st.button("⚠️Consilidar dados", type="primary", use_container_width=True):
                st.warning("⏳ Consolidação iniciada! NÃO feche esta página!")
                
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
