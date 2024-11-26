# measurement.py

import time
import logging
from utils import formatar_tempo, exibir_cabecalho, criar_tabela_temperaturas
from rich.live import Live
from rich.progress import Progress, BarColumn, TimeRemainingColumn
from rich.console import Group
import msvcrt  # Import necessário para detectar teclas no Windows

class Medicao:
    """Controla o processo de medição."""

    def __init__(self, temp_logger, codigos_amostras):
        """Inicializa com o logger de temperatura e códigos das amostras."""
        self.temp_logger = temp_logger  # Thread que lê as temperaturas
        self.codigos_amostras = codigos_amostras  # Códigos das amostras para cada termopar
        self.dados = []  # Lista para armazenar os dados da medição
        self.interrompido = False  # Flag para indicar se a medição foi interrompida

    def realizar_medicao(self, tempos_registro, tempo_total, nome_modo):
        """Realiza a medição de acordo com os tempos especificados."""
        inicio = time.time()
        tempos_registrados = set()  # Conjunto para rastrear os tempos já registrados

        # Inicializa a barra de progresso
        barra_progresso = Progress(
            BarColumn(),
            "[progress.percentage]{task.percentage:>3.0f}%",
            TimeRemainingColumn(),
        )
        tarefa = barra_progresso.add_task("Progresso", total=tempo_total)

        # Inicia o Live rendering
        with Live(refresh_per_second=4, screen=True) as live:
            while True:
                tempo_decorrido = time.time() - inicio
                tempo_atual = int(tempo_decorrido)

                if tempo_atual > tempo_total:
                    break

                # Obter as temperaturas atuais
                temperaturas = self.temp_logger.get_temperaturas()

                # Atualiza a barra de progresso
                barra_progresso.update(tarefa, completed=tempo_atual)

                # Registro de dados nos tempos especificados, apenas uma vez por tempo
                if tempo_atual in tempos_registro and tempo_atual not in tempos_registrados:
                    logging.getLogger().debug(f"Registrando dados no tempo {tempo_atual}: {temperaturas}")
                    self.dados.append((tempo_atual, temperaturas.copy()))
                    tempos_registrados.add(tempo_atual)

                # Atualizar a exibição Live
                self.mostrar_progresso(tempo_atual, tempo_total, temperaturas, barra_progresso, live)

                # Verificar se o usuário pressionou 'M' ou '0'
                if msvcrt.kbhit():
                    tecla = msvcrt.getch().decode('utf-8').upper()
                    if tecla == 'M':
                        # Pausar o Live para interagir com o usuário
                        live.stop()
                        confirmar = input("Deseja interromper a medição e voltar ao menu? (s/n): ")
                        if confirmar.lower() == 's':
                            self.interrompido = True
                            print("Interrompendo a medição e voltando ao menu...")
                            break
                        else:
                            # Retomar o Live
                            live.start()
                    elif tecla == '0':
                        # Pausar o Live para interagir com o usuário
                        live.stop()
                        confirmar = input("Deseja sair do programa? (s/n): ")
                        if confirmar.lower() == 's':
                            self.interrompido = True
                            # Sinalizar para sair do programa
                            raise KeyboardInterrupt("Usuário solicitou saída do programa.")
                        else:
                            # Retomar o Live
                            live.start()

                # Pequena pausa para evitar uso excessivo de CPU
                time.sleep(0.1)

        if not self.interrompido:
            print(f"\n{nome_modo} concluída!")
        return self.dados

    def mostrar_progresso(self, tempo_atual, tempo_total, temperaturas, barra_progresso, live):
        """Atualiza a exibição do progresso e das temperaturas no console."""
        tempo_formatado = formatar_tempo(tempo_atual)

        # Atualizar a tabela de temperaturas
        tabela_temperaturas = criar_tabela_temperaturas(temperaturas, self.codigos_amostras)

        # Mensagem de instruções
        mensagem = "[cyan]Pressione 'M' para interromper e voltar ao menu, '0' para sair do programa.[/cyan]"

        # Criar um grupo de renderizáveis
        grupo_renderizaveis = Group(
            exibir_cabecalho(),
            f"Tempo: {tempo_formatado}",
            tabela_temperaturas,
            barra_progresso,
            mensagem
        )

        # Atualizar o Live com o grupo de renderizáveis
        live.update(grupo_renderizaveis)

    # Funções de medição

    def medir_30s(self):
        """Realiza a medição por 30 segundos, registrando a temperatura em 0s e 30s."""
        tempos_registro = [0, 30]
        tempo_total = 30
        nome_modo = "Análise de amostra comum - 30s"
        return self.realizar_medicao(tempos_registro, tempo_total, nome_modo)

    def medir_2min(self):
        """Realiza a medição por 2 minutos, registrando a temperatura em intervalos de 30s."""
        tempos_registro = [0, 30, 60, 90, 120]
        tempo_total = 120
        nome_modo = "Análise de amostra comum - 2min"
        return self.realizar_medicao(tempos_registro, tempo_total, nome_modo)

    def medir_3min(self):
        """Realiza a medição por 3 minutos, registrando a temperatura em intervalos de 30s."""
        tempos_registro = [0, 30, 60, 90, 120, 150, 180]
        tempo_total = 180
        nome_modo = "Análise de amostra comum - 3min"
        return self.realizar_medicao(tempos_registro, tempo_total, nome_modo)

    def medir_5min(self):
        """Realiza a medição por 5 minutos, registrando a temperatura em intervalos de 1min."""
        tempos_registro = [0] + list(range(60, 301, 60))
        tempo_total = 300
        nome_modo = "Análise de amostra comum - 5min"
        return self.realizar_medicao(tempos_registro, tempo_total, nome_modo)

    def medir_10min(self):
        """Realiza a medição por 10 minutos, registrando a temperatura em intervalos de 1min."""
        tempos_registro = [0] + list(range(60, 601, 60))
        tempo_total = 600
        nome_modo = "Análise de amostra comum - 10min"
        return self.realizar_medicao(tempos_registro, tempo_total, nome_modo)

    def medir_cvmp(self):
        """Realiza a medição específica para Análise CVMP."""
        tempos_registro = [0, 30] + list(range(60, 2101, 60))  # Registro em 0s, 30s e a cada minuto até 35 minutos
        tempo_total = 2100  # 35 minutos
        nome_modo = "Análise CVMP - 30 minutos"
        return self.realizar_medicao(tempos_registro, tempo_total, nome_modo)

    def medir_30min(self):
        """Realiza a medição por 30 minutos, registrando a temperatura em intervalos de 1min."""
        tempos_registro = [0, 30] + list(range(60, 1801, 60))  # Registro em 0s, 30s e a cada minuto até 30 minutos
        tempo_total = 1800
        nome_modo = "Análise de amostra comum - 30min"
        return self.realizar_medicao(tempos_registro, tempo_total, nome_modo)
