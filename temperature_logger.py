# temperature_logger.py

import logging
import threading
import time
import random  # For temperature simulation
from utils import extrair_temperaturas

class SimulatedTemperatureLogger:
    def __init__(self, termopares_ativos):
        self.termopares_ativos = termopares_ativos
        self.temperaturas = {tp: 25.0 for tp in termopares_ativos}  # Start at 25°C
        self.running = False

    def start(self):
        self.running = True
        threading.Thread(target=self.update_temperatures, daemon=True).start()

    def update_temperatures(self):
        max_temperature = 100.0  # Maximum temperature limit
        while self.running:
            for tp in self.termopares_ativos:
                if self.temperaturas[tp] < max_temperature:
                    # Randomly increment the temperature by 0 to 2 degrees
                    self.temperaturas[tp] += random.uniform(0, 2.0)
                    # Ensure it doesn't exceed the maximum
                    if self.temperaturas[tp] > max_temperature:
                        self.temperaturas[tp] = max_temperature
                else:
                    # Optionally, keep the temperature stable or simulate cooling
                    pass  # Do nothing if max temp is reached
            time.sleep(1)  # Update every second

    def parar(self):
        self.running = False

    def get_temperaturas(self):
        return self.temperaturas.copy()

class TemperatureLogger(threading.Thread):
    """Thread that continuously reads temperatures and stores them internally."""

    def __init__(self, manipulador_serial, termopares_ativos, canal_para_termopar):
        """Initializes the thread with the serial handler, active thermocouples, and channel mapping."""
        super().__init__()
        self.manipulador_serial = manipulador_serial
        self.running = True
        self.logger = logging.getLogger("temperature_logger")
        self.termopares_ativos = termopares_ativos
        self.canal_para_termopar = canal_para_termopar
        self.temperaturas = {tp: 'OFF' for tp in termopares_ativos}

    def run(self):
        """Executed when starting the thread."""
        while self.running:
            try:
                dados_brutos = self.manipulador_serial.ler_dados()
                if dados_brutos:
                    dados_decodificados = dados_brutos.decode('utf-8', errors='ignore').strip()
                    self.logger.debug(f"Dados brutos recebidos: {dados_decodificados}")

                    temperaturas = extrair_temperaturas(dados_decodificados, self.canal_para_termopar)
                    if temperaturas:
                        # Updates only the active thermocouples
                        for termopar in self.termopares_ativos:
                            if termopar in temperaturas:
                                temp_value = temperaturas[termopar]
                                # Extract numerical value from temperature string
                                temp_num = float(temp_value.split()[0])
                                self.temperaturas[termopar] = temp_num
                else:
                    self.logger.debug("Nenhum dado recebido. Tentando novamente...")
                    time.sleep(0.1)  # Small pause to avoid excessive CPU usage

            except Exception as e:
                self.logger.error(f"Erro ao ler dados da serial: {e}")
                time.sleep(1)  # Wait before trying again in case of error

    def parar(self):
        """Stops the execution of the thread."""
        self.running = False
        self.logger.info("Thread de leitura de temperatura parada.")

    def get_temperaturas(self):
        """Returns a copy of the current temperatures."""
        return self.temperaturas.copy()
