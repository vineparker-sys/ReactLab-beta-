# main.py

import logging
from logging.handlers import TimedRotatingFileHandler
import time
from serial_handler import ManipuladorPortaSerial
from temperature_logger import TemperatureLogger
from measurement import Medicao
from data_exporter import ExportadorDados
from utils import console, exibir_cabecalho, limpar_tela, extrair_temperaturas
import ctypes
import time
import threading

def impedir_bloqueio(intervalo=60):
    """Impede o bloqueio automático movendo levemente o mouse a cada intervalo de tempo."""
    def mover_mouse():
        while True:
            # Simular um pequeno movimento do mouse
            ctypes.windll.user32.mouse_event(0x0001, 1, 0, 0, 0)  # Move o mouse 1 pixel para a direita
            ctypes.windll.user32.mouse_event(0x0001, -1, 0, 0, 0)  # Move o mouse 1 pixel para a esquerda
            time.sleep(intervalo)  # Espera o intervalo configurado antes de mover novamente

    # Iniciar a thread que impede o bloqueio
    thread = threading.Thread(target=mover_mouse)
    thread.daemon = True  # Thread rodando em segundo plano
    thread.start()

def confirmar_acao(acao):
    """Solicita confirmação do usuário para prosseguir com uma ação."""
    confirmacao = input(f"Tem certeza que deseja {acao}? (s/n): ")
    return confirmacao.lower() == 's'

def mostrar_etapa(etapa):
    """Exibe a etapa atual no console e registra no log."""
    console.print(f"Etapa atual: {etapa}", style="bold yellow")
    logging.info(f"Etapa atual: {etapa}")

import logging
from logging.handlers import TimedRotatingFileHandler
import getpass  # Import necessário para obter o nome do usuário

def configurar_logging():
    """Configura o sistema de logging com diferentes handlers para arquivo e console."""
    # Obter o nome do usuário atual
    username = getpass.getuser()

    # Cria o logger raiz
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)  # Captura todos os níveis de log

    # Formato de log
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # Nome do arquivo de log inclui o nome do usuário
    log_filename = f"debug_{username}.log"

    # Handler para arquivo com rotação (mantém 7 dias de logs)
    file_handler = TimedRotatingFileHandler(log_filename, when="midnight", interval=1, backupCount=7)
    file_handler.setLevel(logging.DEBUG)  # Registrar todos os níveis de log no arquivo
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Handler para console (captura apenas WARNING e superiores)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)  # Exibe apenas WARNING, ERROR e CRITICAL no console
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

def solicitar_codigo_amostras(termopares_ativos):
    """Solicita o código da amostra para cada termopar ativo."""
    codigos_amostras = {}
    for termopar in termopares_ativos:
        while True:
            codigo = input(f"Digite o código da amostra para o {termopar}: ")
            if codigo.lower() == 'menu':
                return "menu"
            elif codigo == 'sair':
                return "sair"
            elif codigo.strip() == '':
                print("Código da amostra não pode estar em branco. Tente novamente.")
            else:
                codigos_amostras[termopar] = codigo.strip()
                break
    return codigos_amostras

def identificar_termopares_ativos(manipulador_serial, canal_para_termopar):
    """Identifica os termopares ativos usando o TemperatureLogger por alguns segundos."""
    logger = logging.getLogger("main")
    logger.info("Iniciando identificação dos termopares ativos...")
    # Inicializar o TemperatureLogger com todos os termopares possíveis
    todos_termopares = ["T1", "T2", "T3", "T4"]
    temp_logger = TemperatureLogger(manipulador_serial, todos_termopares, canal_para_termopar)
    temp_logger.start()
    time.sleep(2)  # Esperar por 2 segundos para coletar dados
    temp_logger.parar()
    temp_logger.join()
    # Verificar quais termopares têm valores diferentes de 'OFF'
    termopares_ativos = [tp for tp, temp in temp_logger.temperaturas.items() if temp != 'OFF']
    logger.info(f"Termopares ativos detectados: {termopares_ativos}")
    return termopares_ativos

def main():
    """Função principal que controla o fluxo do programa."""
    configurar_logging()
    impedir_bloqueio(60)
    programa_rodando = True  # Flag para controlar o loop principal

    while programa_rodando:
        manipulador_serial = None
        temp_logger = None
        dados = None
        try:
            limpar_tela()
            console.print(exibir_cabecalho(), style="bold blue")
            mostrar_etapa("Menu Principal")
            
            print("Escolha uma opção para a análise de reatividade:")
            print("1. 30 segundos")
            print("2. 2 minutos")
            print("3. 3 minutos")
            print("4. 5 minutos")
            print("5. 10 minutos")
            print("6. 30 minutos")
            print("7. 35 minutos")
            print("0. Sair do programa")

            opcao_duracao = input("Digite o número da opção: ")

            if opcao_duracao == '0':
                if confirmar_acao("sair do programa"):
                    print("Saindo do programa...")
                    programa_rodando = False
                    break
                else:
                    continue

            if opcao_duracao not in ['1', '2', '3', '4', '5', '6', '7']:
                print("Opção inválida.")
                continue

            limpar_tela()
            console.print(exibir_cabecalho(), style="bold blue")

            # Mapeamento das opções para funções de medição
            if opcao_duracao == '1':
                funcao_medicao = 'medir_30s'
                descricao_duracao = "[bold red]30 segundos[/bold red]"
                tipo_analise = "comum"
            elif opcao_duracao == '2':
                funcao_medicao = 'medir_2min'
                descricao_duracao = "[bold red]2 minutos[/bold red]"
                tipo_analise = "comum"
            elif opcao_duracao == '3':
                funcao_medicao = 'medir_3min'
                descricao_duracao = "[bold red]3 minutos[/bold red]"
                tipo_analise = "comum"
            elif opcao_duracao == '4':
                funcao_medicao = 'medir_5min'
                descricao_duracao = "[bold red]5 minutos[/bold red]"
                tipo_analise = "comum"
            elif opcao_duracao == '5':
                funcao_medicao = 'medir_10min'
                descricao_duracao = "[bold red]10 minutos[/bold red]"
                tipo_analise = "comum"
            elif opcao_duracao == '6':
                funcao_medicao = 'medir_30min'
                descricao_duracao = "[bold red]30 minutos[/bold red]"
                tipo_analise = "comum"
            elif opcao_duracao == '7':
                funcao_medicao = 'medir_cvmp'
                descricao_duracao = "[bold red]35 minutos[/bold red]"
                tipo_analise = "cvmp"

            # Exibir mensagem da escolha da duração em amarelo com o tempo em vermelho
            console.print(f"[bold yellow]Você escolheu a duração de {descricao_duracao}.[/bold yellow]")
            logging.info(f"Usuário escolheu a duração de {descricao_duracao}.")

            # Abrir a porta serial e lidar com erros
            try:
                porta_com = "COM12"  
                print(f"Tentando abrir a porta serial: {porta_com}")
                logging.info(f"Tentando abrir a porta serial: {porta_com}")
                manipulador_serial = ManipuladorPortaSerial(porta_com)
                if not manipulador_serial.abrir():
                    raise Exception("Erro ao abrir a porta serial. Verifique a conexão.")

                # Definir o mapeamento dos canais para os termopares
                canal_para_termopar = {"41": "T1", "42": "T2", "43": "T3", "44": "T4"}

                # Identificar os termopares ativos
                print("Identificando termopares ativos...")
                termopares_ativos = identificar_termopares_ativos(manipulador_serial, canal_para_termopar)
                print(f"Termopares ativos detectados: {termopares_ativos}")
                if not termopares_ativos:
                    print("Nenhum termopar ativo detectado. Verifique a conexão.")
                    manipulador_serial.fechar()
                    continue  # Volta ao menu principal

                # Solicita códigos das amostras apenas para os termopares ativos
                codigos_amostras = solicitar_codigo_amostras(termopares_ativos)

                if codigos_amostras == "menu":
                    manipulador_serial.fechar()
                    continue  # Volta ao menu principal
                elif codigos_amostras == "sair":
                    manipulador_serial.fechar()
                    print("Saindo do programa...")
                    programa_rodando = False
                    break

                # Exibir a tela de preparação para inserir a amostra e aguardar ENTER
                console.print(
                    "[bold white]Insira a [bold blue]amostra[/bold blue] no recipiente e pressione [bold green]ENTER[/bold green] para iniciar a medição.[/bold white]"
                )
                input()  # Aguarda o usuário pressionar ENTER

                # Inicia a thread de leitura de temperatura com termopares ativos
                logging.info("Iniciando a thread de leitura de temperatura.")
                temp_logger = TemperatureLogger(manipulador_serial, termopares_ativos, canal_para_termopar)
                temp_logger.start()

                # Inicializa a medição com o logger de temperatura e códigos das amostras
                medicao = Medicao(temp_logger, codigos_amostras)

                # Chamar a função de medição específica
                funcao = getattr(medicao, funcao_medicao)
                dados = funcao()

                # Parar a thread de leitura e fechar a porta serial
                if temp_logger:
                    logging.info("Parando a thread de leitura de temperatura.")
                    temp_logger.parar()
                    temp_logger.join()
                if manipulador_serial:
                    manipulador_serial.fechar()

                # Limpar a tela antes de exibir o resultado final
                limpar_tela()

                # Exibir mensagem de conclusão da análise
                console.print(f"[bold blue]Análise de amostra {descricao_duracao} concluída![/bold blue]")

                # Exibir mensagem de fechamento da porta serial
                console.print("[italic grey]Porta serial fechada.[/italic grey]")

                # Perguntar se deseja exportar os dados
                exportar = input("Deseja exportar os dados? (s/n): ")
                if exportar.lower() == 's':
                    exportador = ExportadorDados(dados, codigos_amostras, tipo_analise)
                    caminho_exportacao = exportador.exportar_para_excel()

                    # Limpar a tela após a exportação e exibir apenas o caminho do arquivo
                    limpar_tela()
                    console.print(f"[bold green]Dados exportados com sucesso para a Área de Trabalho![/bold green]")

                    # Perguntar se deseja realizar outra medição
                    nova_medicao = input("Deseja realizar outra medição? (s/n): ")
                    if nova_medicao.lower() != 's':
                        programa_rodando = False
                        break

            except Exception as e:
                # Registrar o erro no arquivo de log
                logging.error(f"Ocorreu um erro: {e}", exc_info=True)
                # Exibir o erro no console de forma destacada
                console.print(f"[bold red]Ocorreu um erro: Erro ao abrir a porta serial. Verifique a conexão.[/bold red]")
                print("\nO que você deseja fazer?")
                print("1. Tentar novamente")
                print("0. Fechar o programa")
                escolha_erro = input("Digite o número da opção: ")

                if escolha_erro == '0':
                    print("Saindo do programa...")
                    programa_rodando = False
                    break
                else:
                    continue

        except KeyboardInterrupt:
            print("\nPrograma encerrado pelo usuário.")
            logging.info("Programa encerrado pelo usuário.")
            programa_rodando = False
            break

        finally:
            # Parar a thread de leitura e fechar a porta serial, se necessário
            if temp_logger and temp_logger.is_alive():
                logging.info("Parando a thread de leitura de temperatura.")
                temp_logger.parar()
                temp_logger.join()
            if manipulador_serial:
                manipulador_serial.fechar()

if __name__ == "__main__":
    main()
