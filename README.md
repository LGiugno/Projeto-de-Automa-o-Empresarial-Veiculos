# Robô Fechamento Veic/Maq - MultiA

Automação de fechamento de laudos de **veículos e máquinas** para os sistemas **MultiA Mais** e **MultiA Avaliações**. Lê dados de uma planilha Google Sheets, envia comparativos via API REST e atualiza automaticamente todos os campos da avaliação — tudo em segundo plano, sem abrir navegador.

---

## Funcionalidades

- Cadastro automático de comparativos com upload de imagens
- Leitura de dados diretamente do Google Sheets (fonte, ano/modelo, KM/horas, valor, OBS)
- Detecção automática de **Máquina vs Veículo** pelo tipo do bem (B13) — altera os campos enviados à API
- Atualização de campos da avaliação: `VLRFIPE`, `REFCONSID`, `FATORDEP`, `FATORVALO`, `FATORCOMERC`, `PERCENTFORCADA`, `OBSLAUDO`, `VALIDADELAUDO`
- Inversão automática do fator de comercialização (lógica do sistema MultiA)
- Ativação do toggle `MEMORIALCALC` após o upload dos comparativos
- Interface gráfica dark mode com log colorido em tempo real
- Cancelamento a qualquer momento

---

## Pré-requisitos

- Python 3.11+
- Google Service Account com acesso à planilha (ver abaixo)

---

## Instalação

```bash
pip install -r requirements.txt
```

---

## Configuração

### 1. Variáveis de ambiente

Copie `.env.example` para `.env` e preencha com suas credenciais:

```bash
cp .env.example .env
```

O `.env` está no `.gitignore` e **nunca deve ser commitado**.

### 2. Credenciais Google

Crie um Service Account no [Google Cloud Console](https://console.cloud.google.com/):

1. Crie um projeto e ative **Google Sheets API** + **Google Drive API**
2. Crie um Service Account e baixe a chave JSON
3. Renomeie para `credentials.json` e coloque na pasta do projeto
4. Compartilhe a planilha com o e-mail do Service Account

Use `credentials.example.json` como referência da estrutura esperada.

---

## Execução

```bash
python fechamento_veic_maq.py
```

---

## Compilar para EXE (Windows)

```bash
pyinstaller --onefile --windowed --icon="logo-fechamento-carros.ico" fechamento_veic_maq.py
```

---

## Estrutura de pastas esperada

```
Comparativos/
├── ABC-1234/        ← nome = placa ou documento
│   ├── 1.jpg
│   └── 2.png
└── DEF-5678/
    ├── 1.png
    ├── 2.png
    └── 3.jpg
```

Cada subpasta deve ter o nome da placa/documento. As imagens devem ser nomeadas com o número do comparativo.

---

## Estrutura da planilha Google Sheets

Cada aba corresponde a um veículo/máquina (nome da aba = nome da subpasta).

| Célula | Conteúdo |
|---|---|
| B4 | Modelo do bem (→ `MODELO`, usado em Máquinas) |
| D5 | Observação do laudo (→ `OBSLAUDO`) |
| B13 | Tipo do bem — se contiver "máquina", ativa modo Máquina |
| K16:K25 | Número do comparativo |
| L16:L25 | Fonte |
| M16:M25 | Ano/modelo (→ `ANOMODELO` ou `FABRICACAO`) |
| N16:N25 | KM ou horas de uso (→ `KM` e `HORASUSO`) |
| O16:O25 | Valor (→ `VALOR`) |
| R16:R25 | Observação do comparativo (→ `OBSERVACOES`, ignorado se vazio) |
| B29 | Valor FIPE (→ `VLRFIPE`) |
| B31 | Referência considerada (→ `REFCONSID`) |
| B34 / D34 | Fator Depreciação / motivo (→ `FATORDEP` / `MOTIVFATORDEP`) |
| B35 / D35 | Fator Valorização / motivo (ignorado se vazio) |
| B36 / D36 | Fator Comercialização / motivo (→ `FATORCOMERC`, **invertido**) |
| B40 | Liquidação Forçada (→ `PERCENTFORCADA`) |

> **Fator de comercialização:** o valor da planilha é automaticamente invertido antes de enviar à API (ex: `10` → `-10`).

---

## Endpoints da API utilizados

| Método | Endpoint | Descrição |
|---|---|---|
| GET | `/multia/avaliacoes` | Busca avaliação por documento/placa |
| GET | `/multia/dadosavaliacao/{uuid}` | Dados completos da avaliação |
| POST | `/multia/adicionarcomparativo/{uuid}` | Adiciona comparativo com imagem |
| POST | `/multia/editaravaliacao/{uuid}` | Atualiza campos da avaliação |
