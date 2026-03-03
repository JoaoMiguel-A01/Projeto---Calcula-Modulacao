# Projeto — Cálculo de Custo de Modulação

Automação em Python para executar diariamente o cálculo do custo de modulação a partir do **PLD horário do submercado SUDESTE**, gerar uma **planilha de modulação** baseada em template e (opcionalmente) **forçar recálculo/salvar/fechar no Excel** via VBScript. citeturn1search1turn11search3turn11search2turn12search2

---

## Visão geral

O fluxo completo é orquestrado por `executar_modulacao.py`, que:
- cria/garante a estrutura de pastas de saída do projeto (PLD, planilha final, logs e gráficos); citeturn1search1
- atualiza automaticamente as rotas internas no `config.ini` com base no diretório onde o projeto está instalado; citeturn1search1turn12search1
- executa a cascata de scripts (download do PLD → geração da planilha → notificação). citeturn1search1

---

## Como funciona (cascata)

### PASSO 1 — Baixar PLD horário e gerar XLSX filtrado
Script: `Src/baixar_pld_ccee_sudeste_xlsx.py` citeturn1search1turn11search3

O que este script faz:
- baixa o CSV do PLD horário via Dados Abertos CCEE (URL definida no `config.ini`); citeturn11search3turn12search1
- filtra por `SUBMERCADO` (padrão: `SUDESTE`), convertendo o resultado para `.xlsx` com `openpyxl`; citeturn11search3turn12search1
- (por padrão) ordena por `MES_REFERENCIA` (desc), `DIA` (desc) e `HORA` (asc); citeturn11search3
- renomeia o arquivo final para um padrão com timestamp: `preco_horario_<submercado> - <YYYY-MM-DDTHHMMSS.mmm>.xlsx`. citeturn11search3

Resiliência do PASSO 1:
- o orquestrador roda o script com tentativas e espera configuráveis (`MAX_TENTATIVAS_CCEE`, `TEMPO_ESPERA_MINUTOS`). citeturn1search1turn12search1

---

### PASSO 2 — Gerar planilha diária a partir do template
Script: `Src/gerar_modulacao_parada_diaria_v3.py` citeturn1search1turn11search2

O que este script faz:
- define a **data-alvo** conforme a regra da janela do PLD (cutoff padrão 17h); citeturn11search2
- localiza o XLSX mais recente do PLD na pasta configurada, valida e carrega os 24 valores (0..23) do dia-alvo; citeturn11search2
- copia o template para a pasta de saída e preenche (exemplo da lógica atual):
  - `B1` = consumo médio do mês; `B2` = total de recurso do mês (ambos vindos do `config.ini`); citeturn11search2turn12search1
  - coluna `A` (linhas 6..29) = dia do mês; coluna `C` (linhas 6..29) = PLD hora a hora (24 valores). citeturn11search2
- gera o arquivo com nome contendo **data + hora (HHMMSS)** para evitar duplicidade. citeturn11search2

Fim de semana (opcional):
- se `USAR_TEMPLATE_FIM_DE_SEMANA=True` e for sábado/domingo, tenta usar o template alternativo `AAAA.MM.DD_Modulacao_Consumo e Cessao - FimDeSemana.xlsx` (se existir). citeturn11search2turn12search1

---

### PASSO 3 — Notificação
O `executar_modulacao.py` chama um script de notificação após o PASSO 2 (no código atual, o alvo é `Src/NotificaCustoModulacao.py`). citeturn1search1

> Observação: se o seu repositório tiver o script de notificação e/ou regras adicionais, documente nesta seção (mensagem enviada, anexos, critérios e erros tratados).

---

### (Opcional) Etapa VBS — Recalcular / Salvar / Fechar no Excel
O gerador da planilha (`gerar_modulacao_parada_diaria_v3.py`) pode executar um VBScript ao final para abrir o XLSX no Excel, forçar o recálculo (via Save) e fechar. citeturn11search2turn12search2

- Script: `Src/recalcular_salvar_fechar.vbs` (caminho configurado no `config.ini`) citeturn12search1turn12search2
- O VBS recebe o caminho do arquivo XLSX como primeiro argumento, abre o Excel invisível, salva e fecha, registrando log `recalcular_salvar.log`. citeturn12search2

---

## Estrutura do projeto

Estrutura mínima esperada: citeturn1search1turn11search3turn11search2

```
Projeto/
├── Configuracoes/
│   └── config.ini
├── Src/
│   ├── baixar_pld_ccee_sudeste_xlsx.py
│   ├── gerar_modulacao_parada_diaria_v3.py
│   └── recalcular_salvar_fechar.vbs
├── Templates/
│   ├── AAAA.MM.DD_Modulacao_Consumo e Cessao.xlsx
│   └── AAAA.MM.DD_Modulacao_Consumo e Cessao - FimDeSemana.xlsx   (opcional)
└── executar_modulacao.py
```

Pastas geradas/garantidas automaticamente pelo orquestrador: `PLD_Horario_Sudeste/`, `Planilha_Modulacao/`, `Logs/`, `GraficosPLD/`. citeturn1search1

---

## Pré-requisitos

- Python 3.8+.
- Dependências (o orquestrador tenta instalar automaticamente na primeira execução): `pandas`, `requests`, `matplotlib`, `openpyxl`. citeturn1search1
- Para a etapa VBS: Windows com Excel disponível (o VBS usa `Excel.Application`). citeturn12search2

---

## Configuração (config.ini)

O projeto lê as configurações em `Configuracoes/config.ini`. citeturn11search3turn11search2turn12search1

### Seções principais

1) `[DIRETORIOS]`
- `PLD_HORARIO`: pasta onde o XLSX do PLD será salvo
- `TEMPLATE_PLANILHA`: caminho do template
- `SAIDA_PLANILHAS`: pasta de saída da planilha final
- `DIRETORIO_LOGS`: pasta de logs
- `VBS_SCRIPT`: caminho do `.vbs` (opcional)
- `TEMP_IMG`: caminho de imagem temporária (se aplicável)

Essas rotas podem ser atualizadas automaticamente pelo orquestrador. citeturn1search1turn12search1

2) `[CCEE_DOWNLOAD]`
- `URL`: endpoint de download (Dados Abertos CCEE)
- `MAX_RETRIES`, `CONNECT_TIMEOUT`, `READ_TIMEOUT`, `USER_AGENT`
- `SUBMERCADO`: submercado a filtrar (padrão: `SUDESTE`)
- `EXPECTED_COLUMNS`, `PRICE_COL_CANDIDATES` (validações e fallback de coluna de preço)

O downloader lê esses valores para baixar e converter o PLD. citeturn11search3turn12search1

3) `[REGRAS_NEGOCIO]`
- `CONSUMO_MEDIO_MWM_MES`: 12 valores (um por mês)
- `TOTAL_RECURSO_MES`: 12 valores (um por mês)
- `CONSUMO_REDUZIDO_MWM`
- `USAR_TEMPLATE_FIM_DE_SEMANA`: True/False

O gerador usa essas listas para preencher o template conforme o mês da data-alvo. citeturn11search2turn12search1

4) `[ORQUESTRADOR]`
- `MAX_TENTATIVAS_CCEE`
- `TEMPO_ESPERA_MINUTOS`

O orquestrador usa isso para controlar tentativas/espera no PASSO 1. citeturn1search1turn12search1

5) `[TELEGRAM]`
- `BOT_TOKEN` e `CHAT_ID` (somente se houver notificação via Telegram)

O orquestrador verifica a presença dessas chaves no início para evitar falhas silenciosas. citeturn1search1turn12search1

### Modelo de `config.ini` (exemplo)

Recomenda-se manter um `config.ini.example` no repositório e não versionar credenciais reais.

```ini
[DIRETORIOS]
PLD_HORARIO = <caminho>\PLD_Horario_Sudeste
TEMPLATE_PLANILHA = <caminho>\Templates\AAAA.MM.DD_Modulacao_Consumo e Cessao.xlsx
SAIDA_PLANILHAS = <caminho>\Planilha_Modulacao
DIRETORIO_LOGS = <caminho>\Logs
VBS_SCRIPT = <caminho>\Src\recalcular_salvar_fechar.vbs
TEMP_IMG = <caminho>\GraficosPLD\grafico_pld.png

[TELEGRAM]
BOT_TOKEN = SEU_TOKEN_AQUI
CHAT_ID = SEU_CHAT_ID_AQUI

[CCEE_DOWNLOAD]
URL = <url_dados_abertos>/content
LOG_NAME = baixar_pld_ccee_sudeste_xlsx.log
MAX_RETRIES = 5
CONNECT_TIMEOUT = 30
READ_TIMEOUT = 120
USER_AGENT = PLD-Downloader/1.1
EXPECTED_COLUMNS = MES_REFERENCIA, SUBMERCADO, PERIODO_COMERCIALIZACAO, DIA, HORA, PLD_HORA
PRICE_COL_CANDIDATES = PLD_HORA, PRECO, PRECO_HORA, PRECO_HORARIO, PRECO_MWH
SUBMERCADO = SUDESTE

[REGRAS_NEGOCIO]
CONSUMO_REDUZIDO_MWM = 4.0
CONSUMO_MEDIO_MWM_MES = 84.0, 90.0, 90.0, 90.0, 90.0, 90.0, 110.0, 125.0, 125.0, 125.0, 125.0, 125.0
TOTAL_RECURSO_MES = 145.37, 149.11, 142.75, 137.50, 129.05, 128.15, 108.20, 122.14, 128.73, 121.22, 117.45, 117.42
USAR_TEMPLATE_FIM_DE_SEMANA = True

[ORQUESTRADOR]
MAX_TENTATIVAS_CCEE = 25
TEMPO_ESPERA_MINUTOS = 30
```

---

## Como executar

### Execução manual

Na raiz do projeto:

```bash
python executar_modulacao.py
```

O orquestrador executa os passos em sequência e encerra com erro se alguma validação falhar. citeturn1search1

---

## Agendamento (Windows Task Scheduler)

Sugestão: agendar para **17:10** (para capturar PLD do dia seguinte quando aplicável). citeturn11search2

Configuração típica:
- Programa: caminho do `python.exe`
- Argumentos: `executar_modulacao.py`
- Iniciar em: pasta raiz do projeto

---

## Saídas geradas

Ao final de uma execução bem-sucedida, você deve encontrar:
- `PLD_Horario_Sudeste/`: arquivo `.xlsx` do PLD filtrado e renomeado com timestamp; citeturn11search3turn1search1
- `Planilha_Modulacao/`: planilha diária gerada a partir do template; citeturn11search2turn1search1
- `Logs/`: logs do orquestrador e scripts; citeturn1search1turn11search3turn11search2
- `Src/recalcular_salvar.log` (ou `%TEMP%\recalcular_salvar.log`): log do VBS quando utilizado. citeturn12search2

---

## Regra do PLD (janela das 17h) — data-alvo

A automação segue a regra de janela:
- se horário local **>= 17:00**, data-alvo = **amanhã**;
- se horário local **< 17:00**, data-alvo = **hoje**. citeturn11search2

Você também pode forçar a data por parâmetro no gerador (`--date YYYY-MM-DD`). citeturn11search2

---

## Formato do XLSX do PLD

O XLSX gerado pelo downloader cria uma aba `PLD_<SUBMERCADO>` (ex.: `PLD_SUDESTE`) e preserva as colunas do CSV do PLD horário. citeturn11search3

Colunas esperadas (exemplo):
- `MES_REFERENCIA`, `SUBMERCADO`, `PERIODO_COMERCIALIZACAO`, `DIA`, `HORA`, `PLD_HORA` citeturn11search3turn12search1

O gerador exige **24 valores** para a data-alvo (horas 0..23); se faltar alguma hora, ele falha. citeturn11search2

---

## Opções úteis de linha de comando

### Downloader
`python Src/baixar_pld_ccee_sudeste_xlsx.py`

Opções comuns:
- `--submercado SUDESTE` (sobrescreve o config.ini) citeturn11search3
- `--keep-csv` (mantém o CSV após converter) citeturn11search3
- `--no-sort` (não ordena) citeturn11search3
- `--no-rename` (não aplica o padrão `preco_horario_<submercado> - ...`) citeturn11search3

### Gerador
`python Src/gerar_modulacao_parada_diaria_v3.py`

Opções comuns:
- `--date YYYY-MM-DD` (força a data-alvo) citeturn11search2
- `--cutoff-hour 17` (altera a hora de corte) citeturn11search2
- `--skip-vbs` (não roda VBS) citeturn11search2
- `--vbs-engine cscript|wscript` e `--vbs-timeout 300` citeturn11search2

---

## Troubleshooting

1) "Ficou tentando baixar PLD por muito tempo"
- ajuste `MAX_TENTATIVAS_CCEE` e `TEMPO_ESPERA_MINUTOS` no `[ORQUESTRADOR]` citeturn1search1turn12search1
- ajuste `MAX_RETRIES` e timeouts no `[CCEE_DOWNLOAD]` citeturn11search3turn12search1

2) "Baixou arquivo, mas não gera a planilha"
- confirme se o XLSX contém a data-alvo completa (horas 0..23); citeturn11search2
- confira se as colunas esperadas estão presentes no XLSX do PLD. citeturn11search2turn11search3

3) "VBS falhou ao abrir Excel"
- verifique se está no Windows e se o Excel está disponível (o VBS usa `Excel.Application`); citeturn12search2
- confira o log `recalcular_salvar.log`. citeturn12search2

4) "Erro de biblioteca (openpyxl etc.)"
- instale manualmente: `pip install openpyxl pandas requests matplotlib` (em ambientes com bloqueio de rede o auto-install pode falhar). citeturn1search1turn11search2turn11search3

---

## Notas

- Os caminhos em `[DIRETORIOS]` podem ser reescritos automaticamente pelo orquestrador ao mover o projeto de pasta/máquina, evitando quebra de rotas. citeturn1search1turn12search1
- O VBS escreve log na pasta do script ou em `%TEMP%` caso não tenha permissão na pasta original. citeturn12search2
