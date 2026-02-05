# FinDash (GitHub Pages) • 2026

Dashboard financeiro moderno, com duas páginas:

- **Dashboard**: usa a **Tabela1** (Entrada/Saída/Líquido/Diferença M-1/Crescimento) + cartões **Nubank** (Tabela2) e **Santander** (Tabela3)
- **Detalhe mensal**: usa os blocos por mês (Tabela4) para listar **Entradas** e **Saídas** com descrições e valores

## Publicar no GitHub Pages

1. Suba todos os arquivos deste ZIP no seu repositório
2. Settings → **Pages**
   - Source: `Deploy from a branch`
   - Branch: `main` / folder `/ (root)`
3. Acesse a URL do GitHub Pages

## Atualização automática (Power Automate → GitHub)

O site NÃO lê o XLSX diretamente. Ele lê este arquivo:

- `data/data.json`

O Power Automate deve sobrescrever esse arquivo sempre que sua planilha do OneDrive mudar.

Veja o guia em `power-automate.md`.

## Formato esperado do data.json

Veja um exemplo real em `data/data.json` (gerado a partir da sua planilha).
