# Preenchimento da planilha padrão — app Streamlit

App web para automatizar o preenchimento da planilha padrão de **Análise por
Máquina** com os arquivos `.xls` baixados dos analisadores de energia — sem
copiar e colar dado por dado.

Suba a planilha modelo (`.xlsm`) e os arquivos do analisador, confira o plano
de preenchimento que o app monta sozinho, e baixe o arquivo pronto.

## Como funciona

A planilha padrão tem 49 abas numeradas ("1" a "49"), uma por dia de medição.
O app:

1. Lê a data já cadastrada na aba **"1"** (célula A6) como referência do ID 1.
2. Para cada arquivo `.xls` enviado, descobre a data do primeiro registro e
   calcula em qual aba (ID) esse dia deve entrar — cada ID a mais é 1 dia a
   mais a partir da referência.
3. Mostra um plano (o que é novo, o que vai ser substituído, o que ficaria
   fora do intervalo 1-49) para você conferir antes de aplicar.
4. Escreve os dados só nas colunas A:M das abas certas, e limpa qualquer
   resíduo de preenchimento antigo que sobre abaixo dos dados novos.
5. Devolve o arquivo `.xlsm` pronto — as fórmulas (Resumo, M1-M7) recalculam
   sozinhas quando você abre no Excel.

O app **nunca mexe em fórmulas**, só nas colunas de dados brutos.

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre em `http://localhost:8501`.

## Publicar no Streamlit Community Cloud (gratuito)

1. Suba este repositório pro seu GitHub (veja abaixo se ainda não tem um).
2. Entre em **[share.streamlit.io](https://share.streamlit.io)** com sua
   conta do GitHub.
3. Clique em **New app**, escolha este repositório, o branch `main` e o
   arquivo `app.py`.
4. Clique em **Deploy** — em 1-2 minutos o app fica no ar com uma URL própria
   (tipo `seu-app.streamlit.app`) que você pode compartilhar com quem quiser.

Não precisa de nenhuma configuração extra — o tema já vem definido em
`.streamlit/config.toml`.

## Criar o repositório no GitHub (se ainda não tiver um)

Eu não consigo criar contas nem repositórios em nome de ninguém — mas é
rápido fazer você mesmo:

1. Entre em [github.com/new](https://github.com/new), dê um nome ao
   repositório (ex.: `preenchimento-planilha-energia`) e clique em
   **Create repository** (pode deixar público ou privado).
2. Na pasta deste projeto, no seu computador, rode:

```bash
git init
git add .
git commit -m "App de preenchimento da planilha padrão"
git branch -M main
git remote add origin https://github.com/SEU-USUARIO/NOME-DO-REPOSITORIO.git
git push -u origin main
```

Pronto — o repositório está no GitHub e já pode ser conectado ao Streamlit
Community Cloud (passo anterior).

## Estrutura do projeto

```
.
├── app.py                     # interface Streamlit
├── preenchimento.py           # lógica de leitura e preenchimento (sem depender do Streamlit)
├── requirements.txt
├── .streamlit/
│   └── config.toml            # tema visual
└── README.md
```

## Limitações conhecidas

- O app não recalcula as fórmulas da planilha (isso pediria instalar o
  LibreOffice no servidor, o que deixaria a hospedagem gratuita mais pesada
  e lenta). Abrir o arquivo no Excel resolve isso automaticamente.
- Feito e testado para o layout específico da planilha padrão descrita
  acima (abas numeradas 1-49, dados em A:M a partir da linha 6). Se o layout
  do seu modelo for diferente, os pontos de ajuste ficam todos em
  `preenchimento.py`.
