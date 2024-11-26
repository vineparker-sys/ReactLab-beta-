# temperature_logger.py

import logging
import threading
import time
from utils import extrair_temperaturas

class TemperatureLogger(threading.Thread):
    """Thread que lê continuamente as temperaturas e as armazena internamente."""

    def __init__(self, manipulador_serial, termopares_ativos, canal_para_termopar):
        """Inicializa a thread com o manipulador serial, termopares ativos e mapeamento de canais."""
        super().__init__()
        self.manipulador_serial = manipulador_serial
        self.running = True
        self.logger = logging.getLogger("temperature_logger")
        self.termopares_ativos = termopares_ativos
        self.canal_para_termopar = canal_para_termopar
        self.temperaturas = {tp: 'OFF' for tp in termopares_ativos}

    def run(self):
        """Executado ao iniciar a thread."""
        while self.running:
            try:
                dados_brutos = self.manipulador_serial.ler_dados()
                if dados_brutos:
                    dados_decodificados = dados_brutos.decode('utf-8', errors='ignore').strip()
                    self.logger.debug(f"Dados brutos recebidos: {dados_decodificados}")

                    temperaturas = extrair_temperaturas(dados_decodificados, self.canal_para_termopar)
                    if temperaturas:
                        # Atualiza apenas os termopares ativos
                        for termopar in self.termopares_ativos:
                            if termopar in temperaturas:
                                self.temperaturas[termopar] = temperaturas[termopar]
                else:
                    self.logger.debug("Nenhum dado recebido. Tentando novamente...")
                    time.sleep(0.1)  # Pequena pausa para evitar uso excessivo de CPU

            except Exception as e:
                self.logger.error(f"Erro ao ler dados da serial: {e}")
                time.sleep(1)  # Esperar antes de tentar novamente em caso de erro

    def parar(self):
        """Interrompe a execução da thread."""
        self.running = False
        self.logger.info("Thread de leitura de temperatura parada.")

    def get_temperaturas(self):
        """Retorna uma cópia das temperaturas atuais."""
        return self.temperaturas.copy()
