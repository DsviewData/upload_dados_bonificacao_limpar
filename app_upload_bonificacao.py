import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
from io import BytesIO
from msal import ConfidentialClientApplication
import unicodedata
import logging
import os
import json
import uuid
import time

# ===========================
# CONFIGURAÇÕES DE VERSÃO - BONIFICAÇÃO v1.0.0
# ===========================
APP_VERSION = "1.0.0"
VERSION_DATE = "2025-10-03"
CHANGELOG = {
    "1.0.0": {
        "date": "2025-10-03",
        "changes": [
            "🎯 Sistema de consolidação de bonificações",
            "🏪 Consolidação por LOJA + MÊS/ANO",
            "📅 Campo DATA_ULTIMO_ENVIO por loja",
            "🎨 Interface moderna e responsiva",
            "🔒 Sistema de lock para múltiplos usuários",
            "💾 Backups automáticos",
            "🛡️ Validação rigorosa de datas",
            "📊 Dashboard com métricas visuais"
        ]
    }
}

# ===========================
# ESTILOS CSS MELHORADOS
# ===========================
def aplicar_estilos_css():
    """Aplica estilos CSS customizados para melhorar a aparência"""
    st.markdown("""
    <style>
    /* Tema principal */
    :root {
        --primary-color: #2E8B57;
        --secondary-color: #20B2AA;
        --accent-color: #FFD700;
        --success-color: #32CD32;
        --warning-color: #FFA500;
        --error-color: #DC143C;
        --background-light: #F8F9FA;
        --text-dark: #2C3E50;
        --border-color: #E1E8ED;
    }
    
    /* Header principal */
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
        font-size: 2.5rem;
        font-weight: 700;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .version-badge {
        background: rgba(255,255,255,0.2);
        padding: 0.5rem 1rem;
        border-radius: 25px;
        font-size: 0.9rem;
        font-weight: 600;
        backdrop-filter: blur(10px);
    }
    
    /* Cards de status */
    .status-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border-left: 4px solid var(--primary-color);
        margin: 1rem 0;
        transition: transform 0.2s ease;
    }
    
    .status-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 20px rgba(0,0,0,0.12);
    }
    
    .status-card.success {
        border-left-color: var(--success-color);
        background: linear-gradient(135deg, #f0fff4, #ffffff);
    }
    
    .status-card.warning {
        border-left-color: var(--warning-color);
        background: linear-gradient(135deg, #fffaf0, #ffffff);
    }
    
    .status-card.error {
        border-left-color: var(--error-color);
        background: linear-gradient(135deg, #fff0f0, #ffffff);
    }
    
    /* Métricas melhoradas */
    .metric-container {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border-top: 3px solid var(--primary-color);
        transition: all 0.3s ease;
    }
    
    .metric-container:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 25px rgba(0,0,0,0.15);
    }
    
    .metric-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: var(--primary-color);
        margin: 0.5rem 0;
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: var(--text-dark);
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Botões melhorados */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
        border: none;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    
    /* Progress bar customizada */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, var(--primary-color), var(--secondary-color));
        border-radius: 10px;
    }
    
    /* Alertas customizados */
    .custom-alert {
        padding: 1rem 1.5rem;
        border-radius: 8px;
        margin: 1rem 0;
        border-left: 4px solid;
        font-weight: 500;
    }
    
    .custom-alert.info {
        background: #e3f2fd;
        border-left-color: #2196f3;
        color: #0d47a1;
    }
    
    .custom-alert.success {
        background: #e8f5e8;
        border-left-color: #4caf50;
        color: #1b5e20;
    }
    
    .custom-alert.warning {
        background: #fff3e0;
        border-left-color: #ff9800;
        color: #e65100;
    }
    
    .custom-alert.error {
        background: #ffebee;
        border-left-color: #f44336;
        color: #b71c1c;
    }
    
    /* Animações */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .fade-in {
        animation: fadeIn 0.6s ease-out;
    }
    
    /* Responsividade */
    @media (max-width: 768px) {
        .main-header h1 {
            font-size: 2rem;
        }
        
        .metric-value {
            font-size: 2rem;
        }
    }
    </style>
    """, unsafe_allow_html=True)

# ===========================
# CONFIGURAÇÃO DE LOGGING
# ===========================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ===========================
# CREDENCIAIS VIA ST.SECRETS
# ===========================
try:
    CLIENT_ID = st.secrets["CLIENT_ID"]
    CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
    TENANT_ID = st.secrets["TENANT_ID"]
    EMAIL_ONEDRIVE = st.secrets["EMAIL_ONEDRIVE"]
    SITE_ID = st.secrets["SITE_ID"]
    DRIVE_ID = st.secrets["DRIVE_ID"]
except KeyError as e:
    st.error(f"❌ Credencial não encontrada: {e}")
    st.stop()

# ===========================
# CONFIGURAÇÃO DE PASTAS
# ===========================
PASTA_CONSOLIDADO = "Documentos Compartilhados/Bonificacao/FonteDeDados"
PASTA_ENVIOS_BACKUPS = "Documentos Compartilhados/PlanilhasEnviadas_Backups/Bonificacao"
PASTA = PASTA_CONSOLIDADO

# ===========================
# CONFIGURAÇÃO DO SISTEMA DE LOCK
# ===========================
ARQUIVO_LOCK = "sistema_lock_bonificacao.json"
TIMEOUT_LOCK_MINUTOS = 10

# ===========================
# AUTENTICAÇÃO
# ===========================
@st.cache_data(ttl=3300)
def obter_token():
    """Obtém token de acesso para Microsoft Graph API"""
    try:
        app = ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" not in result:
            error_desc = result.get("error_description", "Token não obtido")
            st.error(f"❌ Falha na autenticação: {error_desc}")
            return None
            
        return result["access_token"]
        
    except Exception as e:
        st.error(f"❌ Erro na autenticação: {str(e)}")
        logger.error(f"Erro de autenticação: {e}")
        return None

# ===========================
# SISTEMA DE LOCK
# ===========================
def gerar_id_sessao():
    """Gera um ID único para a sessão atual"""
    if 'session_id' not in st.session_state:
        st.session_state.session_id = str(uuid.uuid4())[:8]
    return st.session_state.session_id

def verificar_lock_existente(token):
    """Verifica se existe um lock ativo no sistema"""
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{ARQUIVO_LOCK}:/content"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            lock_data = response.json()
            timestamp_lock = datetime.fromisoformat(lock_data['timestamp'])
            agora = datetime.now()
            
            if agora - timestamp_lock > timedelta(minutes=TIMEOUT_LOCK_MINUTOS):
                logger.info(f"Lock expirado removido automaticamente. Era de {timestamp_lock}")
                remover_lock(token, force=True)
                return False, None
            
            return True, lock_data
        
        elif response.status_code == 404:
            return False, None
        else:
            logger.warning(f"Erro ao verificar lock: {response.status_code}")
            return False, None
            
    except Exception as e:
        logger.error(f"Erro ao verificar lock: {e}")
        return False, None

def criar_lock(token, operacao="Consolidação de bonificações"):
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
        
        response = requests.put(url, headers=headers, data=json.dumps(lock_data))
        
        if response.status_code in [200, 201]:
            logger.info(f"Lock criado com sucesso. Session ID: {session_id}")
            return True, session_id
        else:
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
                logger.warning("Tentativa de remover lock de outra sessão!")
                return False
        
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{ARQUIVO_LOCK}"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.delete(url, headers=headers)
        
        if response.status_code in [200, 204]:
            logger.info("Lock removido com sucesso")
            return True
        elif response.status_code == 404:
            return True
        else:
            logger.error(f"Erro ao remover lock: {response.status_code}")
            return False
            
    except Exception as e:
        logger.error(f"Erro ao remover lock: {e}")
        return False

def atualizar_status_lock(token, session_id, novo_status, detalhes=None):
    """Atualiza o status do lock durante o processo"""
    try:
        lock_existe, lock_data = verificar_lock_existente(token)
        
        if not lock_existe or lock_data.get('session_id') != session_id:
            logger.warning("Lock não existe ou não pertence a esta sessão")
            return False
        
        lock_data['status'] = novo_status
        lock_data['ultima_atualizacao'] = datetime.now().isoformat()
        
        if detalhes:
            lock_data['detalhes'] = detalhes
        
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{ARQUIVO_LOCK}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        response = requests.put(url, headers=headers, data=json.dumps(lock_data))
        return response.status_code in [200, 201]
        
    except Exception as e:
        logger.error(f"Erro ao atualizar status do lock: {e}")
        return False

def exibir_status_sistema(token):
    """Exibe o status atual do sistema de lock"""
    lock_existe, lock_data = verificar_lock_existente(token)
    
    if lock_existe:
        timestamp_inicio = datetime.fromisoformat(lock_data['timestamp'])
        duracao = datetime.now() - timestamp_inicio
        
        st.markdown("""
        <div class="status-card error">
            <h3>🔒 Sistema Ocupado</h3>
            <p>Outro usuário está enviando dados no momento</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-value">{int(duracao.total_seconds()//60)}</div>
                <div class="metric-label">Minutos Ativo</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            tempo_limite = timestamp_inicio + timedelta(minutes=TIMEOUT_LOCK_MINUTOS)
            tempo_restante = tempo_limite - datetime.now()
            minutos_restantes = max(0, int(tempo_restante.total_seconds()//60))
            
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-value">{minutos_restantes}</div>
                <div class="metric-label">Min. Restantes</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            status_display = lock_data.get('status', 'N/A')
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-label">Status</div>
                <div style="font-size: 1.2rem; font-weight: 600; color: var(--warning-color);">{status_display}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with st.expander("ℹ️ Detalhes do processo em andamento"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.info(f"**Operação:** {lock_data.get('operacao', 'N/A')}")
                st.info(f"**Início:** {timestamp_inicio.strftime('%H:%M:%S')}")
                
            with col2:
                if 'detalhes' in lock_data:
                    st.info(f"**Detalhes:** {lock_data['detalhes']}")
                    
                session_id_display = lock_data.get('session_id', 'N/A')[:8]
                st.caption(f"Session ID: {session_id_display}")
        
        if tempo_restante.total_seconds() < 0:
            if st.button("🆘 Liberar Sistema (Forçar)", type="secondary"):
                if remover_lock(token, force=True):
                    st.success("✅ Sistema liberado com sucesso!")
                    st.rerun()
                else:
                    st.error("❌ Erro ao liberar sistema")
        
        return True
    else:
        st.markdown("""
        <div class="status-card success">
            <h3>✅ Sistema Disponível</h3>
            <p>Você pode enviar sua planilha agora</p>
        </div>
        """, unsafe_allow_html=True)
        return False

# ===========================
# FUNÇÕES AUXILIARES
# ===========================
def criar_pasta_se_nao_existir(caminho_pasta, token):
    """Cria pasta no OneDrive se não existir"""
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
            response = requests.get(url, headers=headers)
            
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
                    json=create_body
                )
                
                if create_response.status_code not in [200, 201]:
                    logger.warning(f"Não foi possível criar pasta {parte}")
                    
    except Exception as e:
        logger.warning(f"Erro ao criar estrutura de pastas: {e}")

def upload_onedrive(nome_arquivo, conteudo_arquivo, token, tipo_arquivo="consolidado"):
    """Faz upload de arquivo para OneDrive"""
    try:
        if tipo_arquivo == "consolidado":
            pasta_base = PASTA_CONSOLIDADO
        elif tipo_arquivo in ["enviado", "backup"]:
            pasta_base = PASTA_ENVIOS_BACKUPS
        else:
            pasta_base = PASTA_CONSOLIDADO
        
        pasta_arquivo = "/".join(nome_arquivo.split("/")[:-1]) if "/" in nome_arquivo else ""
        if pasta_arquivo:
            criar_pasta_se_nao_existir(f"{pasta_base}/{pasta_arquivo}", token)
        
        if tipo_arquivo == "consolidado" and "/" not in nome_arquivo:
            mover_arquivo_existente(nome_arquivo, token, pasta_base)
        
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{pasta_base}/{nome_arquivo}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }
        response = requests.put(url, headers=headers, data=conteudo_arquivo)
        
        return response.status_code in [200, 201], response.status_code, response.text
        
    except Exception as e:
        logger.error(f"Erro no upload: {e}")
        return False, 500, f"Erro interno: {str(e)}"

def mover_arquivo_existente(nome_arquivo, token, pasta_base=None):
    """Move arquivo existente para backup antes de substituir"""
    try:
        if pasta_base is None:
            pasta_base = PASTA_CONSOLIDADO
            
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{pasta_base}/{nome_arquivo}"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            file_id = response.json().get("id")
            timestamp = datetime.now().strftime("%Y-%m-%d_%Hh%M")
            nome_base = nome_arquivo.replace(".xlsx", "")
            novo_nome = f"{nome_base}_backup_{timestamp}.xlsx"
            
            patch_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/items/{file_id}"
            patch_body = {"name": novo_nome}
            patch_headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            patch_response = requests.patch(patch_url, headers=patch_headers, json=patch_body)
            
            if patch_response.status_code in [200, 201]:
                st.info(f"💾 Backup criado: {novo_nome}")
            else:
                st.warning(f"⚠️ Não foi possível criar backup do arquivo existente")
                
    except Exception as e:
        st.warning(f"⚠️ Erro ao processar backup: {str(e)}")
        logger.error(f"Erro no backup: {e}")

# ===========================
# VALIDAÇÃO DE DATAS
# ===========================
def validar_datas_detalhadamente(df):
    """Validação detalhada de datas"""
    problemas = []
    
    logger.info(f"🔍 Iniciando validação detalhada de {len(df)} registros...")
    
    for idx, row in df.iterrows():
        linha_excel = idx + 2
        valor_original = row["DATA"]
        loja = row.get("LOJA", "N/A")
        
        problema_encontrado = None
        tipo_problema = None
        
        if pd.isna(valor_original) or str(valor_original).strip() == "":
            problema_encontrado = "Data vazia ou nula"
            tipo_problema = "VAZIO"
        else:
            try:
                data_convertida = pd.to_datetime(valor_original, errors='raise')
                hoje = datetime.now()
                
                if data_convertida > hoje + pd.Timedelta(days=730):
                    problema_encontrado = f"Data muito distante no futuro: {data_convertida.strftime('%d/%m/%Y')}"
                    tipo_problema = "FUTURO"
                elif data_convertida < pd.Timestamp('2020-01-01'):
                    problema_encontrado = f"Data muito antiga: {data_convertida.strftime('%d/%m/%Y')}"
                    tipo_problema = "ANTIGA"
                elif data_convertida > hoje:
                    problema_encontrado = f"Data no futuro: {data_convertida.strftime('%d/%m/%Y')}"
                    tipo_problema = "FUTURO"
                    
            except (ValueError, TypeError, pd.errors.OutOfBoundsDatetime) as e:
                if "day is out of range for month" in str(e) or "month must be in 1..12" in str(e):
                    problema_encontrado = f"Data impossível: {valor_original}"
                    tipo_problema = "IMPOSSÍVEL"
                else:
                    problema_encontrado = f"Formato inválido: {valor_original}"
                    tipo_problema = "FORMATO"
        
        if problema_encontrado:
            problemas.append({
                "Linha Excel": linha_excel,
                "Loja": loja,
                "Valor Original": valor_original,
                "Problema": problema_encontrado,
                "Tipo Problema": tipo_problema
            })
            
            logger.warning(f"❌ Linha {linha_excel}: {problema_encontrado} (Loja: {loja})")
    
    if problemas:
        logger.error(f"❌ TOTAL DE PROBLEMAS ENCONTRADOS: {len(problemas)}")
    else:
        logger.info("✅ Todas as datas estão válidas!")
    
    return problemas

def exibir_problemas_datas(problemas_datas):
    """Exibe problemas de datas com visual melhorado"""
    if not problemas_datas:
        return
    
    st.markdown("""
    <div class="custom-alert error">
        <h4>❌ Problemas de Data Encontrados</h4>
        <p>É obrigatório corrigir TODOS os problemas antes de enviar</p>
    </div>
    """, unsafe_allow_html=True)
    
    tipos_problema = {}
    for problema in problemas_datas:
        tipo = problema["Tipo Problema"]
        tipos_problema[tipo] = tipos_problema.get(tipo, 0) + 1
    
    cols = st.columns(len(tipos_problema))
    emoji_map = {
        "VAZIO": "🔴",
        "FORMATO": "🟠", 
        "IMPOSSÍVEL": "🟣",
        "FUTURO": "🟡",
        "ANTIGA": "🟤"
    }
    
    for i, (tipo, qtd) in enumerate(tipos_problema.items()):
        with cols[i]:
            emoji = emoji_map.get(tipo, "❌")
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-value">{qtd}</div>
                <div class="metric-label">{emoji} {tipo}</div>
            </div>
            """, unsafe_allow_html=True)
    
    df_problemas = pd.DataFrame(problemas_datas)
    max_linhas_exibir = 50
    
    if len(df_problemas) > max_linhas_exibir:
        df_problemas_exibir = df_problemas.head(max_linhas_exibir)
        st.warning(f"⚠️ **Exibindo apenas as primeiras {max_linhas_exibir} linhas.** Total de problemas: {len(df_problemas)}")
    else:
        df_problemas_exibir = df_problemas
    
    st.dataframe(
        df_problemas_exibir,
        use_container_width=True,
        hide_index=True,
        height=400
    )

def validar_dados_enviados(df):
    """Validação super rigorosa dos dados enviados"""
    erros = []
    avisos = []
    linhas_invalidas_detalhes = []
    
    if df.empty:
        erros.append("❌ A planilha está vazia")
        return erros, avisos, linhas_invalidas_detalhes
    
    if "LOJA" not in df.columns:
        erros.append("⚠️ A planilha deve conter uma coluna 'LOJA'")
        avisos.append("📋 Certifique-se de que sua planilha tenha uma coluna chamada 'LOJA'")
    else:
        lojas_validas = df["LOJA"].notna().sum()
        if lojas_validas == 0:
            erros.append("❌ Nenhuma loja válida encontrada na coluna 'LOJA'")
        else:
            lojas_unicas = df["LOJA"].dropna().unique()
            if len(lojas_unicas) > 0:
                avisos.append(f"🏪 Lojas encontradas: {', '.join(map(str, lojas_unicas[:5]))}")
                if len(lojas_unicas) > 5:
                    avisos.append(f"... e mais {len(lojas_unicas) - 5} lojas")
    
    if "DATA" not in df.columns:
        erros.append("⚠️ A planilha deve conter uma coluna 'DATA'")
        avisos.append("📋 Lembre-se: o arquivo deve ter uma aba chamada 'Dados' com as colunas 'DATA' e 'LOJA'")
    else:
        problemas_datas = validar_datas_detalhadamente(df)
        
        if problemas_datas:
            erros.append(f"❌ {len(problemas_datas)} problemas de data encontrados - CONSOLIDAÇÃO BLOQUEADA")
            erros.append("🔧 É OBRIGATÓRIO corrigir TODOS os problemas antes de enviar")
            linhas_invalidas_detalhes = problemas_datas
        else:
            avisos.append("✅ Todas as datas estão válidas e consistentes!")
    
    return erros, avisos, linhas_invalidas_detalhes

# ===========================
# FUNÇÕES DE CONSOLIDAÇÃO
# ===========================
def baixar_arquivo_consolidado(token):
    """Baixa o arquivo consolidado existente"""
    consolidado_nome = "bonificacao_consolidada.xlsx"
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{PASTA_CONSOLIDADO}/{consolidado_nome}:/content"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            df_consolidado = pd.read_excel(BytesIO(response.content))
            df_consolidado.columns = df_consolidado.columns.str.strip().str.upper()
            
            logger.info(f"✅ Arquivo consolidado baixado: {len(df_consolidado)} registros")
            return df_consolidado, True
        else:
            logger.info("📄 Arquivo consolidado não existe - será criado novo")
            return pd.DataFrame(), False
            
    except Exception as e:
        logger.error(f"Erro ao baixar arquivo consolidado: {e}")
        return pd.DataFrame(), False

def adicionar_data_ultimo_envio(df_final, lojas_atualizadas):
    """Adiciona/atualiza a coluna DATA_ULTIMO_ENVIO para as lojas que foram atualizadas"""
    try:
        if 'DATA_ULTIMO_ENVIO' not in df_final.columns:
            df_final['DATA_ULTIMO_ENVIO'] = pd.NaT
            logger.info("➕ Coluna 'DATA_ULTIMO_ENVIO' criada")
        
        data_atual = datetime.now()
        
        for loja in lojas_atualizadas:
            mask = df_final['LOJA'].astype(str).str.strip().str.upper() == str(loja).strip().upper()
            df_final.loc[mask, 'DATA_ULTIMO_ENVIO'] = data_atual
            
            registros_atualizados = mask.sum()
            logger.info(f"📅 Data do último envio atualizada para loja '{loja}': {registros_atualizados} registros")
        
        return df_final
        
    except Exception as e:
        logger.error(f"Erro ao adicionar data do último envio: {e}")
        return df_final

def verificar_seguranca_consolidacao(df_consolidado, df_novo, df_final):
    """Verificação de segurança crítica"""
    try:
        lojas_antes = set(df_consolidado['LOJA'].dropna().astype(str).str.strip().str.upper().unique()) if not df_consolidado.empty else set()
        lojas_novas = set(df_novo['LOJA'].dropna().astype(str).str.strip().str.upper().unique())
        lojas_depois = set(df_final['LOJA'].dropna().astype(str).str.strip().str.upper().unique())
        
        logger.info(f"🛡️ VERIFICAÇÃO DE SEGURANÇA:")
        logger.info(f"   Lojas ANTES: {lojas_antes}")
        logger.info(f"   Lojas NOVAS: {lojas_novas}")
        logger.info(f"   Lojas DEPOIS: {lojas_depois}")
        
        lojas_esperadas = lojas_antes.union(lojas_novas)
        lojas_perdidas = lojas_esperadas - lojas_depois
        
        if lojas_perdidas:
            error_msg = f"Lojas perdidas durante consolidação: {', '.join(lojas_perdidas)}"
            logger.error(f"❌ ERRO CRÍTICO: {error_msg}")
            return False, error_msg
        
        for loja in lojas_antes:
            if loja not in lojas_novas:
                if loja not in lojas_depois:
                    error_msg = f"Loja '{loja}' foi removida sem justificativa"
                    logger.error(f"❌ ERRO: {error_msg}")
                    return False, error_msg
        
        logger.info(f"✅ VERIFICAÇÃO DE SEGURANÇA PASSOU!")
        return True, f"Verificação passou: {len(lojas_depois)} lojas mantidas"
        
    except Exception as e:
        error_msg = f"Erro durante verificação de segurança: {str(e)}"
        logger.error(f"❌ {error_msg}")
        return False, error_msg

def comparar_e_atualizar_registros(df_consolidado, df_novo):
    """Lógica de consolidação por LOJA + MÊS/ANO"""
    registros_inseridos = 0
    registros_substituidos = 0
    registros_removidos = 0
    detalhes_operacao = []
    combinacoes_novas = 0
    combinacoes_existentes = 0
    lojas_atualizadas = set()
    
    logger.info(f"🔧 INICIANDO CONSOLIDAÇÃO:")
    logger.info(f"   Consolidado atual: {len(df_consolidado)} registros")
    logger.info(f"   Novo arquivo: {len(df_novo)} registros")
    
    if df_consolidado.empty:
        df_final = df_novo.copy()
        registros_inseridos = len(df_novo)
        
        lojas_atualizadas = set(df_novo['LOJA'].dropna().astype(str).str.strip().str.upper().unique())
        
        df_temp = df_novo.copy()
        df_temp['mes_ano'] = df_temp['DATA'].dt.to_period('M')
        combinacoes_unicas = df_temp.groupby(['LOJA', 'mes_ano']).size()
        combinacoes_novas = len(combinacoes_unicas)
        
        logger.info(f"✅ PRIMEIRA CONSOLIDAÇÃO: {registros_inseridos} registros inseridos")
        
        for _, row in df_novo.iterrows():
            detalhes_operacao.append({
                "Operação": "INSERIDO",
                "Loja": row["LOJA"],
                "Mês/Ano": row["DATA"].strftime("%m/%Y"),
                "Data": row["DATA"].strftime("%d/%m/%Y"),
                "Motivo": "Primeira consolidação - arquivo vazio"
            })
        
        df_final = adicionar_data_ultimo_envio(df_final, lojas_atualizadas)
        
        return df_final, registros_inseridos, registros_substituidos, registros_removidos, detalhes_operacao, combinacoes_novas, combinacoes_existentes
    
    colunas = df_novo.columns.tolist()
    for col in colunas:
        if col not in df_consolidado.columns:
            df_consolidado[col] = None
            logger.info(f"➕ Coluna '{col}' adicionada ao consolidado")
    
    df_final = df_consolidado.copy()
    
    df_novo_temp = df_novo.copy()
    df_novo_temp['mes_ano'] = df_novo_temp['DATA'].dt.to_period('M')
    
    df_final_temp = df_final.copy()
    df_final_temp['mes_ano'] = df_final_temp['DATA'].dt.to_period('M')
    
    grupos_novos = df_novo_temp.groupby(['LOJA', 'mes_ano'])
    
    logger.info(f"📊 Processando {len(grupos_novos)} combinações únicas de Loja+Mês/Ano")
    
    for (loja, periodo_grupo), grupo_df in grupos_novos:
        if pd.isna(loja) or str(loja).strip() == '':
            logger.warning(f"⚠️ Pulando loja inválida: {loja}")
            continue
        
        lojas_atualizadas.add(str(loja).strip().upper())
        
        logger.info(f"🔍 Processando: '{loja}' em {periodo_grupo} ({len(grupo_df)} registros)")
        
        mask_existente = (
            (df_final_temp["mes_ano"] == periodo_grupo) &
            (df_final_temp["LOJA"].astype(str).str.strip().str.upper() == str(loja).strip().upper())
        )
        
        registros_existentes = df_final[mask_existente]
        
        if not registros_existentes.empty:
            num_removidos = len(registros_existentes)
            
            logger.info(f"   🔄 SUBSTITUIÇÃO: Removendo {num_removidos} registros antigos do período {periodo_grupo}")
            
            df_final = df_final[~mask_existente]
            df_final_temp = df_final_temp[~mask_existente]
            
            registros_removidos += num_removidos
            combinacoes_existentes += 1
            
            detalhes_operacao.append({
                "Operação": "REMOVIDO",
                "Loja": loja,
                "Mês/Ano": periodo_grupo.strftime("%m/%Y"),
                "Data": f"Todo o período {periodo_grupo}",
                "Motivo": f"Substituição: {num_removidos} registro(s) antigo(s) removido(s)"
            })
            
            registros_substituidos += len(grupo_df)
            operacao_tipo = "SUBSTITUÍDO"
            motivo = f"Substituição completa do período: {len(grupo_df)} novo(s) registro(s)"
            
        else:
            logger.info(f"   ➕ NOVA COMBINAÇÃO: Adicionando {len(grupo_df)} registros para {periodo_grupo}")
            registros_inseridos += len(grupo_df)
            combinacoes_novas += 1
            operacao_tipo = "INSERIDO"
            motivo = f"Nova combinação: {len(grupo_df)} registro(s) inserido(s)"
        
        grupo_para_inserir = grupo_df.drop(columns=['mes_ano'], errors='ignore')
        df_final = pd.concat([df_final, grupo_para_inserir], ignore_index=True)
        df_final_temp = pd.concat([df_final_temp, grupo_df], ignore_index=True)
        
        detalhes_operacao.append({
            "Operação": operacao_tipo,
            "Loja": loja,
            "Mês/Ano": periodo_grupo.strftime("%m/%Y"),
            "Data": f"Período {periodo_grupo}",
            "Motivo": motivo
        })
    
    df_final = adicionar_data_ultimo_envio(df_final, lojas_atualizadas)
    
    logger.info(f"🎯 CONSOLIDAÇÃO FINALIZADA:")
    logger.info(f"   Registros inseridos: {registros_inseridos}")
    logger.info(f"   Registros substituídos: {registros_substituidos}")
    logger.info(f"   Registros removidos: {registros_removidos}")
    logger.info(f"   Total final: {len(df_final)} registros")
    
    return df_final, registros_inseridos, registros_substituidos, registros_removidos, detalhes_operacao, combinacoes_novas, combinacoes_existentes

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
        
        sucesso, status_code, resposta = upload_onedrive(nome_arquivo_backup, buffer.read(), token, "backup")
        
        if sucesso:
            logger.info(f"💾 Arquivo enviado salvo como backup: {nome_arquivo_backup}")
        else:
            logger.warning(f"⚠️ Não foi possível salvar backup do arquivo enviado: {status_code}")
            
    except Exception as e:
        logger.error(f"Erro ao salvar arquivo enviado: {e}")

def analise_pre_consolidacao(df_consolidado, df_novo):
    """Análise pré-consolidação"""
    try:
        st.markdown("### 📊 Análise Pré-Consolidação")
        
        df_novo_temp = df_novo.copy()
        df_novo_temp['mes_ano'] = df_novo_temp['DATA'].dt.to_period('M')
        
        lojas_novas = set(df_novo['LOJA'].dropna().astype(str).str.strip().str.upper().unique())
        
        if not df_consolidado.empty:
            df_consolidado_temp = df_consolidado.copy()
            df_consolidado_temp['mes_ano'] = df_consolidado_temp['DATA'].dt.to_period('M')
        
        combinacoes_novas = []
        combinacoes_existentes = []
        
        grupos_novos = df_novo_temp.groupby(['LOJA', 'mes_ano'])
        
        for (loja, periodo), grupo in grupos_novos:
            if pd.isna(loja):
                continue
                
            loja_upper = str(loja).strip().upper()
            
            if not df_consolidado.empty:
                mask_existente = (
                    (df_consolidado_temp["mes_ano"] == periodo) &
                    (df_consolidado_temp["LOJA"].astype(str).str.strip().str.upper() == loja_upper)
                )
                
                if mask_existente.any():
                    combinacoes_existentes.append({
                        "Loja": loja,
                        "Período": periodo.strftime("%m/%Y"),
                        "Novos Registros": len(grupo),
                        "Registros Existentes": mask_existente.sum()
                    })
                else:
                    combinacoes_novas.append({
                        "Loja": loja,
                        "Período": periodo.strftime("%m/%Y"),
                        "Registros": len(grupo)
                    })
            else:
                combinacoes_novas.append({
                    "Loja": loja,
                    "Período": periodo.strftime("%m/%Y"),
                    "Registros": len(grupo)
                })
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-value">{len(lojas_novas)}</div>
                <div class="metric-label">Lojas</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-value">{len(df_novo)}</div>
                <div class="metric-label">Registros Novos</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-value">{len(combinacoes_novas)}</div>
                <div class="metric-label">Novos Períodos</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-value">{len(combinacoes_existentes)}</div>
                <div class="metric-label">Períodos Atualizados</div>
            </div>
            """, unsafe_allow_html=True)
        
        if combinacoes_novas:
            with st.expander("➕ Novos Períodos que serão Adicionados"):
                df_novas = pd.DataFrame(combinacoes_novas)
                st.dataframe(df_novas, use_container_width=True, hide_index=True)
        
        if combinacoes_existentes:
            with st.expander("🔄 Períodos que serão Substituídos"):
                df_existentes = pd.DataFrame(combinacoes_existentes)
                st.dataframe(df_existentes, use_container_width=True, hide_index=True)
        
        return True
        
    except Exception as e:
        logger.error(f"Erro na análise pré-consolidação: {e}")
        st.error(f"❌ Erro na análise: {str(e)}")
        return False

def processar_consolidacao_com_lock(df_novo, nome_arquivo, token):
    """Consolidação com sistema de lock"""
    session_id = gerar_id_sessao()
    
    status_container = st.empty()
    progress_container = st.empty()
    
    try:
        status_container.markdown("""
        <div class="custom-alert info">
            <h4>🔄 Iniciando processo de consolidação...</h4>
        </div>
        """, unsafe_allow_html=True)
        
        sistema_ocupado, lock_data = verificar_lock_existente(token)
        if sistema_ocupado:
            status_container.markdown("""
            <div class="custom-alert error">
                <h4>🔒 Sistema ocupado! Outro usuário está fazendo consolidação.</h4>
            </div>
            """, unsafe_allow_html=True)
            return False
        
        progress_container.progress(10)
        
        lock_criado, session_lock = criar_lock(token, "Consolidação de bonificações")
        
        if not lock_criado:
            status_container.markdown("""
            <div class="custom-alert error">
                <h4>❌ Não foi possível bloquear o sistema. Tente novamente.</h4>
            </div>
            """, unsafe_allow_html=True)
            return False
        
        progress_container.progress(15)
        
        atualizar_status_lock(token, session_lock, "BAIXANDO_ARQUIVO", "Baixando arquivo consolidado")
        progress_container.progress(25)
        
        df_consolidado, arquivo_existe = baixar_arquivo_consolidado(token)
        progress_container.progress(35)

        atualizar_status_lock(token, session_lock, "PREPARANDO_DADOS", "Validando e preparando dados")
        
        df_novo = df_novo.copy()
        df_novo.columns = df_novo.columns.str.strip().str.upper()
        
        df_novo["DATA"] = pd.to_datetime(df_novo["DATA"], errors="coerce")
        linhas_invalidas = df_novo["DATA"].isna().sum()
        df_novo = df_novo.dropna(subset=["DATA"])

        if df_novo.empty:
            status_container.markdown("""
            <div class="custom-alert error">
                <h4>❌ Nenhum registro válido para consolidar</h4>
            </div>
            """, unsafe_allow_html=True)
            remover_lock(token, session_lock)
            return False

        progress_container.progress(45)

        analise_ok = analise_pre_consolidacao(df_consolidado, df_novo)
        
        if not analise_ok:
            remover_lock(token, session_lock)
            return False
        
        progress_container.progress(55)

        atualizar_status_lock(token, session_lock, "CONSOLIDANDO", f"Processando {len(df_novo)} registros por mês/ano")
        progress_container.progress(65)
        
        df_final, inseridos, substituidos, removidos, detalhes, novas_combinacoes, combinacoes_existentes = comparar_e_atualizar_registros(
            df_consolidado, df_novo
        )
        
        progress_container.progress(75)

        verificacao_ok, msg_verificacao = verificar_seguranca_consolidacao(df_consolidado, df_novo, df_final)
        
        if not verificacao_ok:
            status_container.markdown(f"""
            <div class="custom-alert error">
                <h4>❌ ERRO DE SEGURANÇA: {msg_verificacao}</h4>
            </div>
            """, unsafe_allow_html=True)
            remover_lock(token, session_lock)
            return False

        df_final = df_final.sort_values(["DATA", "LOJA"], na_position='last').reset_index(drop=True)
        progress_container.progress(80)
        
        salvar_arquivo_enviado(df_novo, nome_arquivo, token)
        progress_container.progress(85)
        
        atualizar_status_lock(token, session_lock, "UPLOAD_FINAL", "Salvando arquivo consolidado")
        
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name="Dados")
        buffer.seek(0)
        
        consolidado_nome = "bonificacao_consolidada.xlsx"
        sucesso, status_code, resposta = upload_onedrive(consolidado_nome, buffer.read(), token, "consolidado")

        progress_container.progress(95)

        remover_lock(token, session_lock)
        progress_container.progress(100)
        
        if sucesso:
            status_container.empty()
            progress_container.empty()
            
            st.markdown("""
            <div class="custom-alert success">
                <h2>🎉 CONSOLIDAÇÃO REALIZADA COM SUCESSO!</h2>
                <p>🔓 Sistema liberado e disponível para outros usuários</p>
            </div>
            """, unsafe_allow_html=True)
            
            with st.expander("📁 Localização dos Arquivos", expanded=True):
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"📊 **Arquivo Consolidado:**\n`{PASTA_CONSOLIDADO}/bonificacao_consolidada.xlsx`")
                with col2:
                    st.info(f"💾 **Backups e Envios:**\n`{PASTA_ENVIOS_BACKUPS}/`")
            
            st.markdown("### 📈 **Resultado da Consolidação**")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-value">{len(df_final):,}</div>
                    <div class="metric-label">📊 Total Final</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-value">{inseridos}</div>
                    <div class="metric-label">➕ Inseridos</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-value">{substituidos}</div>
                    <div class="metric-label">🔄 Substituídos</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-value">{removidos}</div>
                    <div class="metric-label">🗑️ Removidos</div>
                </div>
                """, unsafe_allow_html=True)
            
            if novas_combinacoes > 0 or combinacoes_existentes > 0:
                st.markdown("### 📈 **Análise de Combinações (Loja + Mês/Ano)**")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown(f"""
                    <div class="metric-container">
                        <div class="metric-value">{novas_combinacoes}</div>
                        <div class="metric-label">🆕 Novos Períodos</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="metric-container">
                        <div class="metric-value">{combinacoes_existentes}</div>
                        <div class="metric-label">🔄 Períodos Atualizados</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    total_processadas = novas_combinacoes + combinacoes_existentes
                    st.markdown(f"""
                    <div class="metric-container">
                        <div class="metric-value">{total_processadas}</div>
                        <div class="metric-label">📊 Total Processado</div>
                    </div>
                    """, unsafe_allow_html=True)
            
            if 'DATA_ULTIMO_ENVIO' in df_final.columns:
                st.markdown("""
                <div class="custom-alert success">
                    <h4>📅 Campo "Data do Último Envio" atualizado!</h4>
                    <p>A planilha consolidada inclui a data do último envio para cada loja</p>
                </div>
                """, unsafe_allow_html=True)
            
            if detalhes:
                with st.expander("📋 Detalhes das Operações", expanded=False):
                    df_detalhes = pd.DataFrame(detalhes)
                    st.dataframe(df_detalhes, use_container_width=True, hide_index=True)
            
            if not df_final.empty:
                resumo_lojas = df_final.groupby("LOJA").agg({
                    "DATA": ["count", "min", "max"]
                }).round(0)
                
                resumo_lojas.columns = ["Total Registros", "Data Inicial", "Data Final"]
                resumo_lojas["Data Inicial"] = pd.to_datetime(resumo_lojas["Data Inicial"]).dt.strftime("%d/%m/%Y")
                resumo_lojas["Data Final"] = pd.to_datetime(resumo_lojas["Data Final"]).dt.strftime("%d/%m/%Y")
                
                if 'DATA_ULTIMO_ENVIO' in df_final.columns:
                    ultimo_envio = df_final.groupby("LOJA")["DATA_ULTIMO_ENVIO"].max()
                    ultimo_envio = ultimo_envio.dt.strftime("%d/%m/%Y %H:%M")
                    resumo_lojas["Último Envio"] = ultimo_envio
                
                with st.expander("🏪 Resumo por Loja"):
                    st.dataframe(resumo_lojas, use_container_width=True)
            
            return True
        else:
            status_container.markdown(f"""
            <div class="custom-alert error">
                <h4>❌ Erro no upload: Status {status_code}</h4>
            </div>
            """, unsafe_allow_html=True)
            return False
            
    except Exception as e:
        logger.error(f"Erro na consolidação: {e}")
        remover_lock(token, session_id, force=True)
        
        status_container.markdown(f"""
        <div class="custom-alert error">
            <h4>❌ Erro durante consolidação: {str(e)}</h4>
        </div>
        """, unsafe_allow_html=True)
        progress_container.empty()
        return False

# ===========================
# INTERFACE STREAMLIT
# ===========================
def main():
    st.set_page_config(
        page_title=f"Sistema de Bonificações v{APP_VERSION}", 
        layout="wide",
        initial_sidebar_state="expanded"
    )

    aplicar_estilos_css()

    st.markdown(f"""
    <div class="main-header fade-in">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <h1>🎁 Sistema de Consolidação de Bonificações</h1>
                <p style="margin: 0.5rem 0 0 0; opacity: 0.9;">Upload e consolidação automática por Loja + Mês/Ano</p>
            </div>
            <div class="version-badge">
                <strong>v{APP_VERSION}</strong><br>
                <small>{VERSION_DATE}</small>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.sidebar.markdown("### 📤 Upload de Bonificações")
    st.sidebar.divider()
    st.sidebar.markdown("**Status do Sistema:**")
    
    token = obter_token()
    if not token:
        st.sidebar.error("❌ Desconectado")
        st.error("❌ Não foi possível autenticar. Verifique as credenciais.")
        st.stop()
    else:
        st.sidebar.success("✅ Conectado")

    st.markdown("## 🔒 Status do Sistema")
    
    sistema_ocupado = exibir_status_sistema(token)
    
    if sistema_ocupado:
        st.markdown("---")
        st.info("🔄 Esta página será atualizada automaticamente a cada 15 segundos")
        time.sleep(15)
        st.rerun()

    st.divider()

    with st.sidebar.expander("ℹ️ Informações do Sistema"):
        st.markdown(f"**Versão:** {APP_VERSION}")
        st.markdown(f"**Data:** {VERSION_DATE}")
        st.markdown(f"**Consolidado:** `bonificacao_consolidada.xlsx`")
        st.markdown(f"**Pasta:** `{PASTA_CONSOLIDADO}`")

    st.markdown("## 📤 Upload de Planilha Excel")
    
    if sistema_ocupado:
        st.markdown("""
        <div class="custom-alert warning">
            <h4>⚠️ Upload desabilitado - Sistema em uso por outro usuário</h4>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🔄 Verificar Status Novamente"):
            st.rerun()
        
        return
    
    st.markdown("""
    <div class="custom-alert info">
        <h4>💡 Importante</h4>
        <p>A planilha deve ter uma aba 'Dados' com as colunas 'LOJA' e 'DATA'</p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.expander("🎯 Como funciona a consolidação", expanded=True):
        st.markdown("### 🏪 **Consolidação por LOJA + MÊS/ANO:**")
        st.info("✅ Substitui dados mensais existentes da mesma loja")
        st.info("✅ Adiciona novos períodos mensais")
        st.info("✅ Mantém dados de outras lojas intactos")
        st.info("✅ Registra data do último envio por loja")
        st.info("✅ Cria backups automáticos")
    
    st.divider()

    uploaded_file = st.file_uploader(
        "Escolha um arquivo Excel", 
        type=["xlsx", "xls"],
        help="Formatos aceitos: .xlsx, .xls"
    )

    df = None
    if uploaded_file:
        try:
            st.markdown(f"""
            <div class="custom-alert success">
                <h4>📁 Arquivo carregado: {uploaded_file.name}</h4>
            </div>
            """, unsafe_allow_html=True)
            
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
                
                with st.expander("👀 Preview dos Dados", expanded=True):
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.markdown(f"""
                        <div class="metric-container">
                            <div class="metric-value">{len(df)}</div>
                            <div class="metric-label">Linhas</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-container">
                            <div class="metric-value">{len(df.columns)}</div>
                            <div class="metric-label">Colunas</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        if "LOJA" in df.columns:
                            lojas_unicas = df["LOJA"].dropna().nunique()
                            st.markdown(f"""
                            <div class="metric-container">
                                <div class="metric-value">{lojas_unicas}</div>
                                <div class="metric-label">Lojas</div>
                            </div>
                            """, unsafe_allow_html=True)
                
        except Exception as e:
            st.error(f"❌ Erro ao ler arquivo: {str(e)}")
            st.stop()

    if df is not None:
        st.markdown("### 🔍 Validação dos Dados")
        
        with st.spinner("🔍 Validando dados..."):
            erros, avisos, problemas_datas = validar_dados_enviados(df)
        
        if erros:
            st.markdown("""
            <div class="custom-alert error">
                <h4>❌ Problemas Encontrados</h4>
            </div>
            """, unsafe_allow_html=True)
            
            for erro in erros:
                st.error(erro)
            
            if problemas_datas:
                exibir_problemas_datas(problemas_datas)
            
            botao_desabilitado = True
        else:
            st.markdown("""
            <div class="custom-alert success">
                <h4>✅ Validação Aprovada</h4>
            </div>
            """, unsafe_allow_html=True)
            botao_desabilitado = False
        
        if avisos:
            for aviso in avisos:
                st.info(aviso)
        
        st.divider()
        
        if not erros:
            col1, col2 = st.columns([2, 1])
            
            with col1:
                if botao_desabilitado:
                    st.button("❌ Consolidar Dados", type="primary", disabled=True)
                else:
                    if st.button("✅ **Consolidar Dados**", type="primary"):
                        st.markdown("""
                        <div class="custom-alert warning">
                            <h4>⏳ Consolidação iniciada! NÃO feche esta página!</h4>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        sucesso = processar_consolidacao_com_lock(df, uploaded_file.name, token)
                        
                        if sucesso:
                            st.balloons()
            
            with col2:
                if st.button("🔄 Limpar Tela", type="secondary"):
                    st.rerun()

    st.markdown("---")
    st.markdown(f"""
    <div style="text-align: center; padding: 2rem; background: #f8f9fa; border-radius: 12px;">
        <strong>🎁 Sistema de Consolidação de Bonificações v{APP_VERSION}</strong><br>
        <small>Última atualização: {VERSION_DATE}</small>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
