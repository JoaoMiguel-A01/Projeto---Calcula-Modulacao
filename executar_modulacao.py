import os
import sys
import subprocess
import configparser
import time
import importlib
from datetime import datetime, timedelta

import openpyxl


# ==========================================
# UTILITÁRIOS
# ==========================================
def imprimir_cabecalho(mensagem):
    print("\n" + "=" * 60)
    print(mensagem)
    print("=" * 60)


def verificar_e_instalar_dependencias(config, caminho_config):
    instaladas = config.getboolean("ORQUESTRADOR", "DEPENDENCIAS_INSTALADAS", fallback=False)
    if instaladas:
        print("\nBibliotecas já instaladas. Avançando...")
        return

    print("\nVerificando bibliotecas necessárias no computador pela primeira vez...")
    dependencias = {
        "pandas": "pandas",
        "requests": "requests",
        "matplotlib": "matplotlib",
        "openpyxl": "openpyxl",
    }

    precisa_reiniciar = False
    for modulo, pacote in dependencias.items():
        try:
            importlib.import_module(modulo)
            print(f" [v] {pacote} já está instalado.")
        except ImportError:
            print(f" [!] Biblioteca '{pacote}' não encontrada. Tentando instalar automaticamente...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pacote], stdout=subprocess.DEVNULL)
                print(f" [+] '{pacote}' instalado com sucesso!")
                precisa_reiniciar = True
            except subprocess.CalledProcessError:
                print(f"\n [ERRO CRÍTICO] Falha ao tentar instalar '{pacote}'.")
                print(" [BLOQUEIO] Sem permissão ou sem Internet para instalar automaticamente.")
                print(f" Instale manualmente:  pip install {pacote}")
                sys.exit(1)

    if "ORQUESTRADOR" not in config:
        config.add_section("ORQUESTRADOR")

    config.set("ORQUESTRADOR", "DEPENDENCIAS_INSTALADAS", "True")
    with open(caminho_config, "w", encoding="utf-8") as configfile:
        config.write(configfile)

    if precisa_reiniciar:
        print(" [!] Novas bibliotecas instaladas. Configuração gravada no config.ini!")


def arquivo_foi_modificado_hoje(pasta, extensao=".xlsx"):
    if not os.path.exists(pasta):
        return False

    hoje = datetime.now().date()
    for arquivo in os.listdir(pasta):
        if arquivo.lower().endswith(extensao):
            caminho = os.path.join(pasta, arquivo)
            if datetime.fromtimestamp(os.path.getmtime(caminho)).date() == hoje:
                return True
    return False


def obter_arquivo_mais_recente(pasta, extensao=".xlsx"):
    if not os.path.exists(pasta):
        return None

    candidatos = [
        os.path.join(pasta, f)
        for f in os.listdir(pasta)
        if f.lower().endswith(extensao)
    ]
    if not candidatos:
        return None

    candidatos.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidatos[0]


# ==========================================
# REGRA DE DATA-ALVO (janela do PLD)
# ==========================================
def calcular_data_alvo_pld():
    """
    - Se horário local >= 17:00 -> data-alvo = amanhã
    - Se horário local <  17:00 -> data-alvo = hoje
    """
    agora = datetime.now()
    if agora.hour >= 17:
        return (agora.date() + timedelta(days=1))
    return agora.date()


# ==========================================
# VALIDAÇÃO DO CONTEÚDO DO PLD (formato real)
# Aba: PLD_SUDESTE
# Colunas esperadas: MES_REFERENCIA, SUBMERCADO, PERIODO_COMERCIALIZACAO, DIA, HORA, PLD_HORA
# ==========================================
def _ler_linha_cabecalho(ws):
    """
    Assume cabeçalho na primeira linha.
    Retorna lista de strings (valores das colunas) normalizados.
    """
    headers = []
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        headers.append(str(v).strip() if v is not None else "")
    return headers


def validar_pld_para_data_alvo(caminho_xlsx, data_alvo, aba="PLD_SUDESTE", exigir_dia_completo=True):
    """
    Valida se o arquivo possui PLD do dia-alvo, conforme o layout observado.

    Checagens:
    1) Aba existe
    2) Cabeçalhos esperados
    3) Existe linhas com MES_REFERENCIA==YYYYMM e DIA==DD e SUBMERCADO==SUDESTE
    4) (opcional) Dia completo: 24 horas (0..23) com PLD_HORA válido
    """
    if not caminho_xlsx or not os.path.exists(caminho_xlsx):
        return False, "Arquivo inexistente."

    try:
        wb = openpyxl.load_workbook(caminho_xlsx, read_only=True, data_only=True)

        if aba not in wb.sheetnames:
            wb.close()
            return False, f"Aba '{aba}' não encontrada no arquivo."

        ws = wb[aba]

        # 1) Cabeçalho
        headers = _ler_linha_cabecalho(ws)
        esperados = ["MES_REFERENCIA", "SUBMERCADO", "PERIODO_COMERCIALIZACAO", "DIA", "HORA", "PLD_HORA"]

        # Verifica se os esperados aparecem na ordem exata (mais rígido e seguro)
        if headers[:6] != esperados:
            wb.close()
            return False, f"Cabeçalho inesperado. Esperado {esperados} e veio {headers[:6]}"

        # 2) Define alvo (YYYYMM e DD)
        mes_ref_alvo = data_alvo.year * 100 + data_alvo.month
        dia_alvo = data_alvo.day

        horas_encontradas = set()
        linhas_dia = 0

        # 3) varre dados a partir da linha 2
        for r in range(2, ws.max_row + 1):
            mes_ref = ws.cell(r, 1).value
            submercado = ws.cell(r, 2).value
            dia = ws.cell(r, 4).value
            hora = ws.cell(r, 5).value
            pld = ws.cell(r, 6).value

            # filtros principais
            if mes_ref is None or dia is None or hora is None:
                continue

            # normalizações simples
            try:
                mes_ref = int(mes_ref)
                dia = int(dia)
                hora = int(hora)
            except Exception:
                continue

            if str(submercado).strip().upper() != "SUDESTE":
                continue

            if mes_ref != mes_ref_alvo or dia != dia_alvo:
                continue

            # achou linha do dia-alvo
            linhas_dia += 1

            # PLD precisa existir e ser numérico
            if pld is None:
                continue
            try:
                float(pld)
            except Exception:
                continue

            horas_encontradas.add(hora)

        wb.close()

        # 4) validações finais
        if linhas_dia == 0:
            return False, f"Não há registros para MES_REFERENCIA={mes_ref_alvo} e DIA={dia_alvo} (SUDESTE)."

        if exigir_dia_completo:
            esperado = set(range(24))
            faltando = sorted(list(esperado - horas_encontradas))
            if faltando:
                return False, f"Dia-alvo encontrado, mas faltam horas: {faltando}"
            return True, "Dia-alvo encontrado com 24 horas válidas."

        return True, "Dia-alvo encontrado (validação simples)."

    except Exception as e:
        return False, f"Erro ao validar planilha: {e}"


def executar_download_pld_com_validacao(script_path, max_tentativas, espera_segundos, pasta_pld):
    """
    Executa o download e valida conteúdo conforme regra da janela do PLD.
    Se não validar, aplica a mesma lógica de quando 'não baixou': espera e tenta novamente.
    """
    data_alvo = calcular_data_alvo_pld()
    print(f"[INFO] Data-alvo do PLD (regra 17h): {data_alvo.strftime('%d/%m/%Y')}")

    for tentativa in range(1, max_tentativas + 1):
        print(f" [Tentativa {tentativa}/{max_tentativas}] Rodando {os.path.basename(script_path)}...")

        try:
            subprocess.run([sys.executable, script_path], check=True)
        except subprocess.CalledProcessError as erro:
            print(f" [AVISO] Script falhou (código {erro.returncode}).")
            if tentativa < max_tentativas:
                print(f" [!] Aguardando {espera_segundos/60:.1f} minutos para tentar novamente...")
                time.sleep(espera_segundos)
                continue
            print(" [ERRO FATAL] Todas as tentativas esgotadas.")
            return False

        # validação 0: arquivo atualizado hoje (mantém sua lógica antiga como pré-filtro)
        if not arquivo_foi_modificado_hoje(pasta_pld):
            print(" [AVISO] Nenhum arquivo .xlsx foi modificado hoje na pasta do PLD.")
        else:
            # validação 1: conteúdo do arquivo mais recente
            arquivo_recente = obter_arquivo_mais_recente(pasta_pld, ".xlsx")
            ok, motivo = validar_pld_para_data_alvo(
                arquivo_recente,
                data_alvo,
                aba="PLD_SUDESTE",
                exigir_dia_completo=True,  # << reforço contra brecha
            )

            if ok:
                print(f" [+] PLD válido confirmado em: {os.path.basename(arquivo_recente)}")
                print(f"     Motivo: {motivo}")
                return True

            print(f" [AVISO] PLD ainda NÃO está válido: {motivo}")

        # mesma lógica de "não baixou"
        if tentativa < max_tentativas:
            print(f" [!] Aguardando {espera_segundos/60:.1f} minutos para tentar novamente...")
            time.sleep(espera_segundos)
        else:
            print(" [ERRO FATAL] Todas as tentativas esgotadas (PLD não validou a data-alvo).")
            return False


# ==========================================
# MAIN
# ==========================================
def main():
    imprimir_cabecalho("INICIANDO MAIN MODULAÇÃO")

    base_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Diretório base: {base_dir}")

    config_path = os.path.join(base_dir, "Configuracoes", "config.ini")
    template_path = os.path.join(base_dir, "Templates", "AAAA.MM.DD_Modulacao_Consumo e Cessao.xlsx")
    vbs_path = os.path.join(base_dir, "Src", "recalcular_salvar_fechar.vbs")

    if not os.path.exists(config_path):
        print(f"\n[BLOQUEIO] Arquivo 'config.ini' não encontrado em: {config_path}")
        sys.exit(1)

    config = configparser.ConfigParser()
    config.optionxform = str
    config.read(config_path, encoding="utf-8")

    if "ORQUESTRADOR" not in config:
        config.add_section("ORQUESTRADOR")

    if not config.has_option("ORQUESTRADOR", "MAX_TENTATIVAS_CCEE"):
        config.set("ORQUESTRADOR", "MAX_TENTATIVAS_CCEE", "20")

    if not config.has_option("ORQUESTRADOR", "TEMPO_ESPERA_MINUTOS"):
        config.set("ORQUESTRADOR", "TEMPO_ESPERA_MINUTOS", "20")

    max_tentativas_ccee = config.getint("ORQUESTRADOR", "MAX_TENTATIVAS_CCEE")
    tempo_espera_minutos = config.getint("ORQUESTRADOR", "TEMPO_ESPERA_MINUTOS")

    verificar_e_instalar_dependencias(config, config_path)

    # Pastas de saída
    pastas_necessarias = ["PLD_Horario_Sudeste", "Planilha_Modulacao", "Logs", "GraficosPLD"]
    print("\nVerificando pastas de saída...")
    for pasta in pastas_necessarias:
        caminho = os.path.join(base_dir, pasta)
        if not os.path.exists(caminho):
            os.makedirs(caminho)
            print(f" [+] Criado: {pasta}")

    # Template
    if not os.path.exists(template_path):
        print("\n[BLOQUEIO] Planilha Template não encontrada na pasta 'Templates'!")
        sys.exit(1)

    # Atualiza INI com rotas dinâmicas
    print("\nAtualizando caminhos internos...")
    if "DIRETORIOS" not in config:
        config.add_section("DIRETORIOS")

    config.set("DIRETORIOS", "PLD_HORARIO", os.path.join(base_dir, "PLD_Horario_Sudeste"))
    config.set("DIRETORIOS", "TEMPLATE_PLANILHA", template_path)
    config.set("DIRETORIOS", "SAIDA_PLANILHAS", os.path.join(base_dir, "Planilha_Modulacao"))
    config.set("DIRETORIOS", "DIRETORIO_LOGS", os.path.join(base_dir, "Logs"))
    config.set("DIRETORIOS", "DIRETORIO_NOTIFICACAO", base_dir)
    config.set("DIRETORIOS", "VBS_SCRIPT", vbs_path)
    config.set("DIRETORIOS", "TEMP_IMG", os.path.join(base_dir, "GraficosPLD", "grafico_pld.png"))

    with open(config_path, "w", encoding="utf-8") as configfile:
        config.write(configfile)

    # Variáveis de negócio obrigatórias
    print("\nVerificando variáveis de negócio...")
    chaves = [
        ("TELEGRAM", "BOT_TOKEN"),
        ("TELEGRAM", "CHAT_ID"),
        ("REGRAS_NEGOCIO", "CONSUMO_MEDIO_MWM_MES"),
        ("REGRAS_NEGOCIO", "TOTAL_RECURSO_MES"),
        ("REGRAS_NEGOCIO", "CONSUMO_REDUZIDO_MWM"),
    ]

    vazios = [
        f"[{secao}] -> {chave}"
        for secao, chave in chaves
        if (not config.has_option(secao, chave)) or (not config.get(secao, chave).strip())
    ]

    if vazios:
        print(" [ERRO CRÍTICO] Preencha no 'config.ini':", ", ".join(vazios))
        sys.exit(1)

    print("\n" + "=" * 60 + "\nINICIANDO CASCATA\n" + "=" * 60)

    # PASSO 1: Baixar PLD com validação robusta
    print(f"\n⏳ PASSO 1: Baixar PLD (Até {max_tentativas_ccee} tentativas)")
    script_pld = os.path.join(base_dir, "Src", "baixar_pld_ccee_sudeste_xlsx.py")
    pasta_pld = os.path.join(base_dir, "PLD_Horario_Sudeste")

    if not executar_download_pld_com_validacao(
        script_pld,
        max_tentativas_ccee,
        tempo_espera_minutos * 60,
        pasta_pld,
    ):
        sys.exit(1)

    # PASSO 2
    print("\n⏳ PASSO 2: Preencher Template")
    script_gerar = os.path.join(base_dir, "Src", "gerar_modulacao_parada_diaria_v3.py")
    subprocess.run([sys.executable, script_gerar], check=True)

    if not arquivo_foi_modificado_hoje(os.path.join(base_dir, "Planilha_Modulacao")):
        sys.exit(1)

    # PASSO 3
    print("\n⏳ PASSO 3: Analisar Cenários e Enviar Telegram")
    script_notifica = os.path.join(base_dir, "Src", "NotificaCustoModulacao.py")
    subprocess.run([sys.executable, script_notifica], check=True)

    imprimir_cabecalho("PROCESSO CONCLUIDO COM SUCESSO!")


if __name__ == "__main__":
    main()