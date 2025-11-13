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
from dateutil.relativedelta import relativedelta

# ===========================
# CONFIGURAÇÕES DE VERSÃO
# ===========================
APP_VERSION = "3.0.0"
VERSION_DATE = "2025-11-13"
APP_TITLE = "Upload da planilha de Bonificações"
APP_SUBTITLE = "Consolidação inteligente por loja e mês"

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
    
    .warning-box {
        background: #fff3cd;
        border-left: 4px solid var(--warning-color);
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
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

def criar_lock(token, operacao="Consolidação por loja e mês"):
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
        
        content = json.dumps(lock_data).encode('utf-8')
        response = requests.put(url, headers=headers, data=content, timeout=10)
        
        if response.status_code in [200, 201]:
            logger.info(f"Lock criado: {session_id}")
            return True
        
        logger.error(f"Falha ao criar lock: {response.status_code}")
        return False
        
    except Exception as e:
        logger.error(f"Erro ao criar lock: {e}")
        return False

def remover_lock(token, session_id=None, force=False):
    """Remove o lock do sistema"""
    try:
        if not force:
            ocupado, lock_data = verificar_lock_existente(token)
            if ocupado and lock_data:
                if session_id and lock_data.get("session_id") != session_id:
                    logger.warning("Tentativa de remover lock de outra sessão")
                    return False
        
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{ARQUIVO_LOCK}"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.delete(url, headers=headers, timeout=10)
        
        if response.status_code in [204, 404]:
            logger.info("Lock removido com sucesso")
            return True
        
        logger.error(f"Falha ao remover lock: {response.status_code}")
        return False
        
    except Exception as e:
        logger.error(f"Erro ao remover lock: {e}")
        return False

def exibir_status_sistema(token):
    """Exibe o status atual do sistema e retorna se está ocupado"""
    ocupado, lock_data = verificar_lock_existente(token)
    
    if ocupado and lock_data:
        st.markdown('<div class="status-card error">', unsafe_allow_html=True)
        st.error("🔒 **Sistema em uso** - Aguarde a conclusão da operação atual")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Operação", lock_data.get('operacao', 'N/A'))
        with col2:
            timestamp = datetime.fromisoformat(lock_data['timestamp'])
            tempo_decorrido = datetime.now() - timestamp
            minutos = int(tempo_decorrido.total_seconds() / 60)
            st.metric("Tempo", f"{minutos} min")
        with col3:
            st.metric("Sessão", lock_data.get('session_id', 'N/A'))
        
        st.markdown('</div>', unsafe_allow_html=True)
        return True
    else:
        st.markdown('<div class="status-card success">', unsafe_allow_html=True)
        st.success("✅ **Sistema disponível** - Pronto para processar")
        st.markdown('</div>', unsafe_allow_html=True)
        return False

# ===========================
# VALIDAÇÃO DE DATAS
# ===========================
def validar_datas(df):
    """
    Valida as datas da planilha
    Retorna: (sucesso, erros, avisos, info)
    """
    erros = []
    avisos = []
    info = {}
    
    # Verifica se tem coluna DATA
    if 'DATA' not in df.columns:
        erros.append("Coluna DATA não encontrada na planilha")
        return False, erros, avisos, info
    
    # Converter para datetime se necessário
    try:
        df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce')
    except Exception as e:
        erros.append(f"Erro ao converter coluna DATA: {str(e)}")
        return False, erros, avisos, info
    
    # Verificar datas nulas
    datas_nulas = df['DATA'].isna().sum()
    if datas_nulas > 0:
        erros.append(f"Encontradas {datas_nulas} linhas com DATA vazia ou inválida")
    
    # Verificar se há datas válidas
    if df['DATA'].notna().sum() == 0:
        erros.append("Nenhuma data válida encontrada na planilha")
        return False, erros, avisos, info
    
    # Data atual e limites
    data_atual = datetime.now()
    mes_atual = data_atual.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    mes_anterior = (mes_atual - relativedelta(months=1))
    mes_proximo = (mes_atual + relativedelta(months=1))
    limite_futuro = (mes_atual + relativedelta(months=2))  # Aceita até 2 meses no futuro
    limite_passado = (mes_atual - relativedelta(months=6))  # Aceita até 6 meses no passado
    
    # Filtrar apenas datas válidas para análise
    datas_validas = df[df['DATA'].notna()]['DATA']
    
    # Verificar datas muito futuras
    datas_futuras = datas_validas[datas_validas > limite_futuro]
    if len(datas_futuras) > 0:
        datas_exemplo = datas_futuras.head(5).dt.strftime('%d/%m/%Y').tolist()
        erros.append(f"⚠️ Encontradas {len(datas_futuras)} datas muito futuras (mais de 2 meses). Exemplos: {', '.join(datas_exemplo)}")
    
    # Verificar datas muito antigas
    datas_antigas = datas_validas[datas_validas < limite_passado]
    if len(datas_antigas) > 0:
        datas_exemplo = datas_antigas.head(5).dt.strftime('%d/%m/%Y').tolist()
        avisos.append(f"⚠️ Encontradas {len(datas_antigas)} datas antigas (mais de 6 meses atrás). Exemplos: {', '.join(datas_exemplo)}")
    
    # Identificar meses presentes nos dados
    df['MES_ANO'] = df['DATA'].dt.to_period('M')
    meses_unicos = df[df['DATA'].notna()]['MES_ANO'].unique()
    
    info['meses_presentes'] = sorted([str(m) for m in meses_unicos])
    info['total_meses'] = len(meses_unicos)
    info['data_minima'] = datas_validas.min()
    info['data_maxima'] = datas_validas.max()
    info['mes_atual'] = str(mes_atual.strftime('%Y-%m'))
    
    # Avisos sobre meses
    if len(meses_unicos) > 1:
        avisos.append(f"📅 Dados contêm {len(meses_unicos)} meses diferentes: {', '.join(info['meses_presentes'])}")
    
    # Verificar se tem dados do mês atual ou anterior
    tem_mes_atual = any(df['MES_ANO'] == mes_atual.strftime('%Y-%m'))
    tem_mes_anterior = any(df['MES_ANO'] == mes_anterior.strftime('%Y-%m'))
    
    if tem_mes_atual:
        avisos.append(f"✅ Dados contêm informações do mês atual ({mes_atual.strftime('%m/%Y')})")
    if tem_mes_anterior:
        avisos.append(f"✅ Dados contêm informações do mês anterior ({mes_anterior.strftime('%m/%Y')})")
    
    # Remover coluna auxiliar antes de retornar
    df.drop('MES_ANO', axis=1, inplace=True, errors='ignore')
    
    sucesso = len(erros) == 0
    return sucesso, erros, avisos, info

# ===========================
# VALIDAÇÃO DE ESTRUTURA
# ===========================
def validar_estrutura_colunas(df, token):
    """Valida a estrutura de colunas do DataFrame"""
    erros = []
    avisos = []
    info = {}
    
    colunas_usuario = [col for col in df.columns if col != 'DATA_ULTIMO_ENVIO']
    colunas_faltando = [col for col in COLUNAS_OBRIGATORIAS if col not in df.columns]
    colunas_novas = [col for col in colunas_usuario if col not in COLUNAS_OBRIGATORIAS]
    
    info['colunas_usuario'] = colunas_usuario
    info['colunas_faltando'] = colunas_faltando
    info['colunas_novas'] = colunas_novas
    
    if colunas_faltando:
        erros.append(f"❌ Colunas obrigatórias ausentes: {', '.join(colunas_faltando)}")
    
    if colunas_novas:
        avisos.append(f"ℹ️ Novas colunas detectadas: {', '.join(colunas_novas)}")
    
    return erros, avisos, info

# ===========================
# VALIDAÇÃO COMPLETA
# ===========================
def validar_dados_enviados(df, token):
    """Validação completa dos dados enviados"""
    erros_totais = []
    avisos_totais = []
    
    # 1. Validar estrutura de colunas
    erros_estrutura, avisos_estrutura, info_estrutura = validar_estrutura_colunas(df, token)
    erros_totais.extend(erros_estrutura)
    avisos_totais.extend(avisos_estrutura)
    
    # 2. Validar datas
    sucesso_datas, erros_datas, avisos_datas, info_datas = validar_datas(df)
    erros_totais.extend(erros_datas)
    avisos_totais.extend(avisos_datas)
    
    # Guardar info de datas no session_state
    if 'info_datas' not in st.session_state:
        st.session_state.info_datas = {}
    st.session_state.info_datas = info_datas
    
    # 3. Validar LOJA
    if 'LOJA' in df.columns:
        lojas_nulas = df['LOJA'].isna().sum()
        if lojas_nulas > 0:
            erros_totais.append(f"❌ Encontradas {lojas_nulas} linhas sem LOJA definida")
        
        lojas_unicas = df['LOJA'].dropna().unique()
        avisos_totais.append(f"📍 Dados contêm {len(lojas_unicas)} lojas diferentes")
    else:
        erros_totais.append("❌ Coluna LOJA não encontrada")
    
    return erros_totais, avisos_totais

# ===========================
# DOWNLOAD DE ARQUIVO
# ===========================
def download_arquivo_sharepoint(token, nome_arquivo):
    """Faz download de um arquivo do SharePoint"""
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{nome_arquivo}:/content"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            return BytesIO(response.content)
        elif response.status_code == 404:
            logger.warning(f"Arquivo não encontrado: {nome_arquivo}")
            return None
        else:
            logger.error(f"Erro ao baixar arquivo: {response.status_code}")
            return None
            
    except Exception as e:
        logger.error(f"Erro no download: {e}")
        return None

# ===========================
# UPLOAD DE ARQUIVO
# ===========================
def upload_arquivo_sharepoint(token, nome_arquivo, conteudo, pasta):
    """Faz upload de um arquivo para o SharePoint"""
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{pasta}/{nome_arquivo}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        
        response = requests.put(url, headers=headers, data=conteudo, timeout=60)
        
        if response.status_code in [200, 201]:
            logger.info(f"Arquivo enviado: {nome_arquivo}")
            return True
        else:
            logger.error(f"Erro no upload: {response.status_code} - {response.text}")
            return False
            
    except Exception as e:
        logger.error(f"Erro no upload: {e}")
        return False

# ===========================
# CONSOLIDAÇÃO INTELIGENTE
# ===========================
def processar_consolidacao_inteligente(df_novo, nome_arquivo_original, token):
    """
    Processa a consolidação inteligente:
    - Identifica lojas e meses nos novos dados
    - Remove registros da mesma loja E mês do consolidado
    - Adiciona os novos registros
    - Preserva todos os outros dados
    """
    session_id = gerar_id_sessao()
    
    try:
        # Criar lock
        st.info("🔒 Bloqueando sistema para consolidação...")
        if not criar_lock(token, "Consolidação por loja e mês"):
            st.error("❌ Não foi possível bloquear o sistema. Tente novamente.")
            return False
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 1. Download do arquivo consolidado
        status_text.info("📥 Baixando arquivo consolidado...")
        progress_bar.progress(10)
        
        arquivo_consolidado = download_arquivo_sharepoint(token, "bonificacao_consolidada.xlsx")
        
        if arquivo_consolidado is None:
            status_text.warning("⚠️ Arquivo consolidado não existe. Criando novo arquivo...")
            df_consolidado = pd.DataFrame()
        else:
            df_consolidado = pd.read_excel(arquivo_consolidado, sheet_name="Dados")
            df_consolidado.columns = df_consolidado.columns.str.strip().str.upper()
            status_text.success(f"✅ Arquivo consolidado carregado: {len(df_consolidado)} registros")
        
        progress_bar.progress(20)
        
        # 2. Preparar dados novos
        status_text.info("🔄 Preparando novos dados...")
        df_novo_processado = df_novo.copy()
        
        # Adicionar DATA_ULTIMO_ENVIO
        df_novo_processado['DATA_ULTIMO_ENVIO'] = datetime.now()
        
        # Garantir que DATA está em datetime
        df_novo_processado['DATA'] = pd.to_datetime(df_novo_processado['DATA'])
        
        # Criar coluna MES_ANO para identificação
        df_novo_processado['MES_ANO'] = df_novo_processado['DATA'].dt.to_period('M').astype(str)
        
        progress_bar.progress(30)
        
        # 3. Identificar lojas e meses nos novos dados
        status_text.info("🔍 Identificando lojas e meses a serem atualizados...")
        
        lojas_meses_novos = df_novo_processado[['LOJA', 'MES_ANO']].drop_duplicates()
        
        total_combinacoes = len(lojas_meses_novos)
        st.info(f"📊 Serão atualizados dados de {total_combinacoes} combinações de loja/mês")
        
        # Exibir detalhes
        with st.expander("📋 Detalhes das atualizações", expanded=True):
            summary = df_novo_processado.groupby(['LOJA', 'MES_ANO']).size().reset_index(name='Quantidade')
            st.dataframe(summary, use_container_width=True)
        
        progress_bar.progress(40)
        
        # 4. Remover registros antigos das mesmas lojas/meses
        if len(df_consolidado) > 0:
            status_text.info("🗑️ Removendo registros antigos das mesmas lojas/meses...")
            
            # Garantir que consolidado também tem MES_ANO
            df_consolidado['DATA'] = pd.to_datetime(df_consolidado['DATA'])
            df_consolidado['MES_ANO'] = df_consolidado['DATA'].dt.to_period('M').astype(str)
            
            registros_antes = len(df_consolidado)
            
            # Criar condição para manter apenas registros que NÃO estão sendo atualizados
            condicao_manter = True
            for _, row in lojas_meses_novos.iterrows():
                loja = row['LOJA']
                mes_ano = row['MES_ANO']
                condicao_manter = condicao_manter & ~((df_consolidado['LOJA'] == loja) & (df_consolidado['MES_ANO'] == mes_ano))
            
            df_consolidado_filtrado = df_consolidado[condicao_manter].copy()
            registros_removidos = registros_antes - len(df_consolidado_filtrado)
            
            st.success(f"✅ {registros_removidos} registros antigos removidos")
            st.info(f"📊 {len(df_consolidado_filtrado)} registros preservados de outros meses/lojas")
        else:
            df_consolidado_filtrado = pd.DataFrame()
            st.info("ℹ️ Não há dados consolidados anteriores")
        
        progress_bar.progress(60)
        
        # 5. Combinar dados
        status_text.info("🔄 Combinando dados...")
        
        # Remover coluna auxiliar MES_ANO antes de consolidar
        df_novo_processado.drop('MES_ANO', axis=1, inplace=True, errors='ignore')
        if len(df_consolidado_filtrado) > 0:
            df_consolidado_filtrado.drop('MES_ANO', axis=1, inplace=True, errors='ignore')
        
        if len(df_consolidado_filtrado) > 0:
            df_final = pd.concat([df_consolidado_filtrado, df_novo_processado], ignore_index=True)
        else:
            df_final = df_novo_processado
        
        st.success(f"✅ Consolidação concluída: {len(df_final)} registros totais")
        
        progress_bar.progress(70)
        
        # 6. Criar backup do arquivo anterior (se existir)
        if arquivo_consolidado is not None:
            status_text.info("💾 Criando backup do arquivo anterior...")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_backup = f"BACKUP_bonificacao_{timestamp}.xlsx"
            
            if upload_arquivo_sharepoint(token, nome_backup, arquivo_consolidado.getvalue(), PASTA_ENVIOS_BACKUPS):
                st.success(f"✅ Backup criado: {nome_backup}")
            else:
                st.warning("⚠️ Não foi possível criar backup, mas continuando...")
        
        progress_bar.progress(80)
        
        # 7. Salvar arquivo consolidado atualizado
        status_text.info("💾 Salvando arquivo consolidado atualizado...")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Dados', index=False)
        output.seek(0)
        
        if not upload_arquivo_sharepoint(token, "bonificacao_consolidada.xlsx", output.getvalue(), PASTA_CONSOLIDADO):
            status_text.error("❌ Erro ao salvar arquivo consolidado")
            remover_lock(token, session_id, force=True)
            return False
        
        st.success("✅ Arquivo consolidado atualizado com sucesso!")
        
        progress_bar.progress(90)
        
        # 8. Salvar cópia do arquivo enviado
        status_text.info("💾 Salvando cópia do arquivo enviado...")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_copia = f"ENVIO_{timestamp}_{nome_arquivo_original}"
        
        output_copia = BytesIO()
        with pd.ExcelWriter(output_copia, engine='openpyxl') as writer:
            df_novo.to_excel(writer, sheet_name='Dados', index=False)
        output_copia.seek(0)
        
        if upload_arquivo_sharepoint(token, nome_copia, output_copia.getvalue(), PASTA_ENVIOS_BACKUPS):
            st.success(f"✅ Cópia salva: {nome_copia}")
        
        progress_bar.progress(95)
        
        # 9. Remover lock
        status_text.info("🔓 Liberando sistema...")
        remover_lock(token, session_id)
        
        progress_bar.progress(100)
        status_text.success("✅ Processo concluído com sucesso!")
        
        # Exibir resumo final
        st.markdown("---")
        st.markdown("### 📊 Resumo da Consolidação")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Registros Novos", len(df_novo))
        with col2:
            st.metric("Registros Removidos", registros_removidos if 'registros_removidos' in locals() else 0)
        with col3:
            st.metric("Registros Preservados", len(df_consolidado_filtrado) if len(df_consolidado_filtrado) > 0 else 0)
        with col4:
            st.metric("Total Final", len(df_final))
        
        return True
        
    except Exception as e:
        logger.error(f"Erro na consolidação: {e}")
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
    st.markdown('<div class="warning-box">', unsafe_allow_html=True)
    st.markdown("""
    ### ⚠️ ATENÇÃO - Nova Lógica de Consolidação
    
    **Este sistema agora funciona de forma INTELIGENTE:**
    
    ✅ **Consolidação Seletiva por Loja e Mês**
    - Identifica a LOJA e o MÊS de cada registro enviado
    - Remove APENAS os registros do consolidado que têm a mesma LOJA e mesmo MÊS
    - Adiciona os novos registros
    - **PRESERVA todos os outros dados** (outras lojas ou outros meses)
    
    ✅ **Validação de Datas**
    - Verifica se as datas estão em formato válido
    - Alerta sobre datas muito futuras (mais de 2 meses à frente)
    - Alerta sobre datas muito antigas (mais de 6 meses atrás)
    
    📌 **Exemplo prático:**
    Se você enviar dados da "LOJA A" referente a "Janeiro/2025", o sistema irá:
    1. Remover apenas os registros antigos da "LOJA A" de "Janeiro/2025"
    2. Manter todos os dados da "LOJA A" de outros meses
    3. Manter todos os dados de outras lojas
    4. Adicionar os novos dados enviados
    """)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.info("""**✨ Funcionalidades:**
- ✅ Validação completa de datas
- ✅ Consolidação inteligente por loja e mês
- ✅ Preserva dados de outros meses e lojas
- ✅ Valida estrutura de colunas obrigatórias
- ✅ Permite adição de novas colunas
- ✅ Faz backup automático antes de consolidar
- ✅ Adiciona campo DATA_ULTIMO_ENVIO em todos os registros
- ✅ Salva cópia do arquivo enviado
    """)

    # Informações do sistema
    with st.sidebar.expander("ℹ️ Informações"):
        st.markdown(f"**Modo:** Consolidação Inteligente")
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
        
        # Mostrar informações sobre os meses que serão atualizados
        if 'info_datas' in st.session_state and st.session_state.info_datas:
            info_datas = st.session_state.info_datas
            
            st.markdown("---")
            st.markdown("### 📅 Resumo dos Dados a Serem Consolidados")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Meses Diferentes", info_datas.get('total_meses', 0))
            with col2:
                if 'data_minima' in info_datas:
                    st.metric("Data Mínima", info_datas['data_minima'].strftime('%d/%m/%Y'))
            with col3:
                if 'data_maxima' in info_datas:
                    st.metric("Data Máxima", info_datas['data_maxima'].strftime('%d/%m/%Y'))
            
            if 'meses_presentes' in info_datas:
                st.info(f"**Meses identificados:** {', '.join(info_datas['meses_presentes'])}")
        
        st.divider()
        
        # Botões de ação
        col1, col2 = st.columns([2, 1])
        
        with col1:
            if st.button("🔄 Consolidar Dados (Inteligente)", type="primary", use_container_width=True):
                st.warning("⏳ Consolidação iniciada! NÃO feche esta página!")
                
                sucesso = processar_consolidacao_inteligente(df, uploaded_file.name, token)
                
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
        <small>Modo: Consolidação Inteligente | Última atualização: {VERSION_DATE}</small>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
