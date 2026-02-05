# Power Automate → GitHub (auto atualização)

Objetivo: sempre que a planilha do OneDrive mudar, atualizar **data/data.json** no seu repositório.

## Passo 1 — Trigger (OneDrive)

- **When a file is modified** (OneDrive)
- Escolha seu arquivo `.xlsx`

## Passo 2 — Ler as tabelas do Excel

Seu arquivo tem 4 partes lógicas (como nas imagens):

1) **Tabela1 (Dashboard)** — A1:F12  
2) **Tabela2 (Nubank)** — H1:I13  
3) **Tabela3 (Santander)** — K1:L13  
4) **Tabela4 (Detalhe mensal)** — blocos por mês começando em N1

✅ Como o Power Automate “Excel Online” funciona melhor com *Tabela* (Insert → Table),
o jeito mais estável é criar **uma tabela única** chamada `Lancamentos`
com as colunas:

- `date` (YYYY-MM-01)
- `type` (`income` ou `expense`)
- `description`
- `amount`
- `account` (opcional: Nubank/Santander)
- `category` (opcional)

**Mas** como sua planilha atual é “layout de blocos”, a forma prática é:

- Você manter a planilha como está
- E criar uma aba/tabela auxiliar `Export` (automática) que “desdobra” os blocos em linhas
- O Power Automate lê apenas a `Export` (List rows present in a table)

### Aba `Export` (recomendado)
Crie uma nova planilha/aba chamada **Export** e monte uma Tabela (Inserir → Tabela) com colunas:

- `month` (Fevereiro, Março, ...)
- `kind` (Entrada/Saída)
- `descricao`
- `valor`
- `account` (opcional)

Essa tabela pode ser preenchida manualmente ou por fórmulas.

## Passo 3 — Montar JSON

No Power Automate:

1. **List rows present in a table** (na aba Export)
2. Use **Data Operations → Select** para mapear para o formato:

```json
{ "descricao": "...", "valor": 123.45 }
```

3. Use um **Compose** para construir o arquivo completo no formato do site:

- `year`: 2026
- `months`: array com Fevereiro → Dezembro
- `summary`: da Tabela1 (pode ser preenchida via outra tabela auxiliar ou manualmente)

👉 Se você preferir, eu posso te entregar uma versão do Excel já com a aba `Export` pronta e fórmulas, para o Power Automate ficar plug-and-play.

## Passo 4 — GitHub: Create or update file

- Repo: seu repositório
- Path: `data/data.json`
- Content: saída do Compose
- Commit message: `Atualiza data.json`

## Observações

- O GitHub Pages pode cachear: no site existe botão **Atualizar agora**.
- Se quiser, configure o fluxo para rodar também a cada X horas (fallback).
