# Projeto — Cálculo de Custo de Modulação

Automação em Python para executar o cálculo do custo de modulação a partir do **PLD horário do submercado SUDESTE**, gerar uma **planilha de modulação** baseada em template, forçar recálculo no Excel via VBScript e enviar um relatório via Telegram.

O sistema opera agora como um **Serviço Contínuo (Daemon)**, não dependendo de agendadores externos para funcionar.

---

## Visão geral

O fluxo completo é orquestrado de forma contínua por `executar_modulacao.py`, que:
- cria/garante a estrutura de pastas de saída do projeto (PLD, planilha final, logs e gráficos);
- atualiza automaticamente as rotas internas no `config.ini` com base no diretório onde o projeto está instalado;
- aguarda silenciosamente o horário agendado (`HORARIO_EXECUCAO`) e executa a cascata de scripts (download do PLD → geração da planilha → notificação) de forma autônoma.

---

## Como funciona (cascata)

### PASSO 1 — Baixar PLD horário e gerar XLSX filtrado
Script: `Src/baixar_pld_ccee_sudeste_xlsx.py`

O que este script faz:
- baixa o CSV do PLD horário via Dados Abertos CCEE (URL definida no `config.ini`);
- filtra por `SUBMERCADO` (padrão: `SUDESTE`), convertendo o resultado para `.xlsx` com `openpyxl`;
- ordena por `MES_REFERENCIA` (desc), `DIA` (desc) e `HORA` (asc);
- renomeia o arquivo final para um padrão com timestamp: `preco_horario_<submercado> - <YYYY-MM-DDTHHMMSS.mmm>.xlsx`.

Resiliência do PASSO 1:
- O orquestrador valida ativamente se os dados do dia correto já constam no XLSX. Se o site estiver fora do ar ou a CCEE não tiver atualizado o dado, o script não trava. Ele aguarda e tenta novamente com base em `MAX_TENTATIVAS_CCEE` e `TEMPO_ESPERA_MINUTOS`.

---

### PASSO 2 — Gerar planilha diária a partir do template
Script: `Src/gerar_modulacao_parada_diaria_v3.py`

O que este script faz:
- define a **data-alvo** conforme a regra da janela do PLD (cutoff padrão 17h);
- localiza o XLSX mais recente do PLD na pasta configurada, valida e carrega os 24 valores (0..23) do dia-alvo;
- copia o template para a pasta de saída e preenche:
  - `B1` = consumo médio do mês; `B2` = total de recurso do mês (ambos vindos do `config.ini`);
  - coluna `A` (linhas 6..29) = dia do mês; coluna `C` (linhas 6..29) = PLD hora a hora (24 valores).
- gera o arquivo com nome contendo **data + hora (HHMMSS)** para evitar duplicidade.

Fim de semana (inteligente):
- se `USAR_TEMPLATE_FIM_DE_SEMANA=True` e for sábado/domingo, tenta usar o template alternativo `AAAA.MM.DD_Modulacao_Consumo e Cessao - FimDeSemana.xlsx`. Se o arquivo não existir, ele avisa e usa o padrão automaticamente para não quebrar a rotina.

---

### PASSO 3 — Notificação e Análise Financeira
Script: `Src/NotificaCustoModulacao.py`

O que este script faz:
- lê a planilha gerada no Passo 2 e calcula cenários de redução (Flat, Redução de 3h e Redução de 4h);
- atualiza fisicamente a planilha pintando as 4 horas mais caras de vermelho e reduzindo o consumo (Coluna E);
- gera um gráfico da variação do PLD do dia;
- aciona o VBScript para recalcular o Excel;
- envia uma mensagem formatada, o gráfico gerado e a planilha `.xlsx` como anexo no Telegram.

---

### Etapa VBS — Recalcular / Salvar / Fechar no Excel
Script: `Src/recalcular_salvar_fechar.vbs`

O VBS é chamado pelo Python para abrir o Excel de forma invisível, forçar o cálculo das fórmulas embutidas na planilha, salvar e fechar. Isso garante que o gestor que abrir o anexo no celular ou PC veja os resultados financeiros atualizados.

---

## Estrutura do projeto

Estrutura mínima esperada antes do primeiro uso:
