# Guia de Deploy

## Preparação

### 1. Criar Repositório GitHub

```bash
# Inicialize o repositório
git init
git add .
git commit -m "Initial commit: Sistema de Bonificações v1.0.0"

# Conecte ao GitHub
git remote add origin https://github.com/seu-usuario/sistema-bonificacao.git
git branch -M main
git push -u origin main
```

### 2. Estrutura Final do Projeto

```
sistema-bonificacao/
├── app_upload_bonificacao_consolidado.py
├── requirements.txt
├── README.md
├── .gitignore
├── DEPLOY.md
├── .streamlit/
│   ├── config.toml
│   └── secrets.toml.example
└── docs/
    └── images/ (opcional - screenshots)
```

## Deploy no Streamlit Cloud

### Passo 1: Acesse Streamlit Cloud
- Vá para [share.streamlit.io](https://share.streamlit.io)
- Faça login com sua conta GitHub

### Passo 2: Novo App
1. Clique em "New app"
2. Selecione seu repositório
3. Branch: `main`
4. Main file path: `app_upload_bonificacao_consolidado.py`

### Passo 3: Configure Secrets
1. Clique em "Advanced settings"
2. Cole o conteúdo do seu `secrets.toml` em "Secrets"
3. Formato:

```toml
CLIENT_ID = "valor-real"
CLIENT_SECRET = "valor-real"
TENANT_ID = "valor-real"
EMAIL_ONEDRIVE = "email-real"
SITE_ID = "valor-real"
DRIVE_ID = "valor-real"
```

### Passo 4: Deploy
- Clique em "Deploy!"
- Aguarde o build (2-5 minutos)
- Sua app estará disponível em: `https://seu-app.streamlit.app`

## Atualizações

Para atualizar a aplicação:

```bash
# Faça suas alterações
git add .
git commit -m "Descrição das alterações"
git push origin main
```

O Streamlit Cloud fará o redeploy automático!

## Troubleshooting

### Erro de Autenticação
- Verifique se todas as secrets estão corretas
- Confirme permissões no Azure AD

### Erro de Timeout
- Aumente `maxUploadSize` no config.toml se necessário

### App não inicia
- Verifique os logs no Streamlit Cloud
- Confirme que requirements.txt está correto

## Monitoramento

- Logs: Disponíveis no painel do Streamlit Cloud
- Métricas: Analytics integrado
- Alertas: Configure notificações por email

## Boas Práticas

1. **Nunca commite secrets.toml**
2. **Use .gitignore apropriado**
3. **Mantenha requirements.txt atualizado**
4. **Documente mudanças no README**
5. **Use versionamento semântico**
6. **Teste localmente antes de fazer push**

## Backup

Recomenda-se:
- Backup regular do código (GitHub)
- Backup dos arquivos consolidados no OneDrive
- Documentação de configurações

## Segurança

- Secrets apenas no Streamlit Cloud
- HTTPS automático
- Validação de dados
- Sistema de lock para concorrência

## Suporte

- Documentação Streamlit: [docs.streamlit.io](https://docs.streamlit.io)
- Microsoft Graph: [docs.microsoft.com/graph](https://docs.microsoft.com/graph)
