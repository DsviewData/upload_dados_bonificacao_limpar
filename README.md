# Sistema de Consolidação de Bonificações

Sistema automatizado para consolidação de planilhas de bonificações por Loja + Mês/Ano.

## Funcionalidades

- Upload de planilhas Excel (.xlsx, .xls)
- Consolidação automática por LOJA + MÊS/ANO
- Sistema de lock para múltiplos usuários
- Validação rigorosa de datas
- Backups automáticos
- Campo DATA_ULTIMO_ENVIO por loja
- Interface moderna e responsiva
- Dashboard com métricas visuais

## Tecnologias

- Python 3.9+
- Streamlit
- Pandas
- Microsoft Graph API (OneDrive/SharePoint)
- MSAL (Microsoft Authentication Library)

## Instalação Local

```bash
# Clone o repositório
git clone https://github.com/seu-usuario/sistema-bonificacao.git
cd sistema-bonificacao

# Crie um ambiente virtual
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

# Instale as dependências
pip install -r requirements.txt

# Configure as credenciais
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Edite secrets.toml com suas credenciais

# Execute a aplicação
streamlit run app_upload_bonificacao_consolidado.py
```

## Configuração

### Credenciais Microsoft Graph API

Crie um arquivo `.streamlit/secrets.toml` com:

```toml
CLIENT_ID = "seu-client-id"
CLIENT_SECRET = "seu-client-secret"
TENANT_ID = "seu-tenant-id"
EMAIL_ONEDRIVE = "email@empresa.com"
SITE_ID = "seu-site-id"
DRIVE_ID = "seu-drive-id"
```

### Como obter as credenciais

1. **Azure AD App Registration:**
   - Acesse [Azure Portal](https://portal.azure.com)
   - Azure Active Directory > App registrations > New registration
   - Copie CLIENT_ID e TENANT_ID
   - Em "Certificates & secrets", crie um CLIENT_SECRET

2. **Permissões necessárias:**
   - Sites.ReadWrite.All
   - Files.ReadWrite.All

3. **Site ID e Drive ID:**
   - Use Microsoft Graph Explorer para obter os IDs

## Deploy no Streamlit Cloud

1. Faça push do código para GitHub
2. Acesse [Streamlit Cloud](https://share.streamlit.io)
3. Conecte seu repositório
4. Adicione as secrets nas configurações do app
5. Deploy!

## Estrutura de Pastas OneDrive

```
Documentos Compartilhados/
├── Bonificacao/
│   └── FonteDeDados/
│       └── bonificacao_consolidada.xlsx
└── PlanilhasEnviadas_Backups/
    └── Bonificacao/
        └── [arquivos enviados com timestamp]
```

## Como Usar

1. Acesse a aplicação
2. Faça upload da planilha Excel
3. A planilha deve ter:
   - Aba chamada "Dados"
   - Coluna "LOJA"
   - Coluna "DATA"
4. Clique em "Consolidar Dados"
5. Aguarde o processamento
6. Verifique o resultado!

## Formato da Planilha

| LOJA | DATA | [outras colunas] |
|------|------|------------------|
| 001  | 01/01/2025 | ... |
| 002  | 01/01/2025 | ... |

## Lógica de Consolidação

- Agrupa por **LOJA + MÊS/ANO**
- Substitui dados mensais existentes da mesma loja
- Adiciona novos períodos mensais
- Mantém dados de outras lojas intactos
- Registra data do último envio por loja

## Sistema de Lock

- Bloqueia o sistema durante consolidação
- Timeout de 10 minutos
- Permite forçar liberação se necessário

## Segurança

- Verificação de dados antes da consolidação
- Backups automáticos antes de substituir
- Validação rigorosa de datas
- Proteção contra perda de dados

## Suporte

Para problemas ou dúvidas, abra uma issue no GitHub.

## Versão

**v1.0.0** - 2025-10-03

## Licença

MIT License
