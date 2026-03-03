=========================================================
AUTOMAÇÃO DE CÁLCULO CUSTO DE MODULAÇÃO - LIASA
=========================================================

1. FUNÇÃO DO SISTEMA
Este orquestrador (executar_modulacao.py) automatiza o processo diário de cálculo de modulação:
- Baixa os dados do PLD Horário Sudeste diretamente do portal da CCEE.
- Cria a estrutura de pastas necessária na máquina local.
- Preenche a planilha template com os dados do PLD e os consumos parametrizados.
- Executa um script VBScript para forçar o recálculo das fórmulas no Excel.
- Analisa os cenários (Flat, Redução 3h, Redução 4h), formata visualmente a planilha e envia os resultados financeiros e o arquivo via Telegram.

2. ESTRUTURA DE PASTAS (Como o projeto deve estar organizado)
Para o sistema funcionar, a pasta principal deve conter a seguinte estrutura base:

Projeto - Modulacao/
├── Configuracoes/     -> Contém o arquvio config.ini
├── Src/               -> Contém os códigos Python e o script .vbs
├── Templates/         -> Contém as planilhas template
└── executar_modulacao.py -> Responsável por gerenciar a execução de todos os outros códigos Python

3. PRÉ-REQUISITOS (Primeira Execução)
- Ter o Python instalado na máquina (versão 3.8 ou superior).
- NÃO é necessário instalar bibliotecas manualmente. O script "executar_modulacao.py" verificará e instalará o Pandas, Requests, Matplotlib e Openpyxl automaticamente na primeira execução.
- Ter a planilha "AAAA.MM.DD_Modulacao_Consumo e Cessao.xlsx" dentro da pasta "Templates".
- (Opcional) Ter a planilha "AAAA.MM.DD_Modulacao_Consumo e Cessao - FimDeSemana.xlsx" na pasta "Templates" caso deseje usar a curva específica para sábados e domingos.

4. COMO CONFIGURAR
1. Extraia esta pasta num diretório à sua escolha, sendo local ou rede.
2. Abra a pasta "Configuracoes" e edite o arquivo "config.ini" com o Bloco de Notas.
3. Preencha as chaves obrigatórias que estão vazias:
   -> [TELEGRAM]: BOT_TOKEN e CHAT_ID
   -> [REGRAS_NEGOCIO]: CONSUMO_MEDIO_MWM_MES, TOTAL_RECURSO_MES, CONSUMO_REDUZIDO_MWM
4. Configure a regra de Fim de Semana:
   -> [REGRAS_NEGOCIO]: USAR_TEMPLATE_FIM_DE_SEMANA = True (Ou mude para False se quiser desativar esta função).
5. Salve e feche o "config.ini". 
(Nota: Não precisa preencher os caminhos das pastas, o sistema fará isso sozinho).

5. COMO EXECUTAR
Basta dar um duplo clique no arquivo "executar_modulacao.py" na raiz do projeto ou agendá-lo no Task Scheduler do Windows. 

O "executar_modulacao.py" tratará de todo o processo de forma sequencial.

6. TRATAMENTO DE ERROS (O que o sistema faz sozinho?)
Para garantir que os cálculos sejam precisos e evitar falhas silenciosas, o sistema foi programado para contornar ou alertar sobre os seguintes cenários:

* Instabilidade no site da CCEE: Se o portal estiver fora do ar ou o download falhar, o script não cancela a operação imediatamente. 
    Ele entra em modo de espera e tenta novamente, repetindo até N vezes (valores ajustáveis no config.ini).
* Validação em Cascata: O sistema é amarrado. Ele só preenche a planilha se garantir que baixou o PLD do dia correto. 
    E só envia a mensagem no Telegram se a planilha física tiver sido gerada e atualizada com sucesso.
* Cenários de Fim de Semana e Fallback: Se a função de fim de semana estiver ativada, mas o arquivo do template de fim de semana for apagado ou não existir na pasta "Templates", 
    o sistema deteta o erro, avisa no terminal e utiliza o template padrão automaticamente para garantir que o cálculo do dia é entregue.
* Bloqueio por falta de dados: Se você esquecer de preencher o Token do Telegram ou alguma variável financeira no config.ini, 
    o sistema realiza um "Pre-flight check" no primeiro segundo de execução. Ele aborta o processo imediatamente e avisa no terminal exatamente qual a informação que está faltando.
* Adaptação a novos computadores: Se você mover a pasta inteira do projeto para outro computador ou disco (ex: de C:\ para D:\), 
    o sistema deteta a mudança e reescreve as rotas internas do config.ini sozinho, sem quebrar os caminhos.
* Falha na instalação de bibliotecas: Se o computador tiver bloqueios de rede que impeçam a auto-instalação das bibliotecas Python, 
    o sistema pausa com segurança e fornece no terminal o comando exato que a equipe de TI precisa rodar manualmente.
