# Arquivos mínimos para atualizar automaticamente via OneDrive -> GitHub

## O que isso faz
- Um GitHub Action roda a cada 30 minutos
- Baixa sua planilha XLSX do OneDrive (URL de download direto)
- Converte para `data/data.json`
- Faz commit e push no repositório

## Passo 1 — Adicionar estes arquivos no seu repo
- `.github/workflows/sync-onedrive.yml`
- `scripts/sync_onedrive_to_datajson.mjs`
- `package.json`

## Passo 2 — Configurar o Secret com a URL do XLSX
No GitHub:
Settings → Secrets and variables → Actions → **New repository secret**
- Name: `ONEDRIVE_XLSX_URL`
- Value: **URL de DOWNLOAD direto do XLSX**

⚠️ Seu link atual é do tipo `.../doc.aspx?...` (visualização).
Ele normalmente NÃO serve para automação, porque retorna HTML.
Você precisa de um link que devolva o arquivo `.xlsx` (download), sem login.

## Passo 3 — Rodar o workflow
Aba Actions → “Sync OneDrive XLSX -> data/data.json” → Run workflow

## Frequência
No arquivo YAML você pode ajustar o `cron`.
