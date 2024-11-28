# measurement.py

import threading
import time
import logging

class Medicao:
    def __init__(self, temperature_logger, codigos_amostras, tempos_registro, tempo_total, nome_modo):
        self.temp_logger = temperature_logger
        self.codigos_amostras = codigos_amostras
        self.tempos_registro = tempos_registro
        self.tempo_total = tempo_total
        self.nome_modo = nome_modo

    def obter_temperaturas(self):
        # Retrieve temperatures from the temperature logger
        return self.temp_logger.get_temperaturas()

    def run(self):
        """Executes the measurement process."""
        self.inicio = time.time()
        tempos_registrados = set()

        while not self.interrompido and (time.time() - self.inicio) <= self.tempo_total:
            tempo_decorrido = time.time() - self.inicio
            tempo_atual = int(tempo_decorrido)

            # Get current temperatures
            self.temperaturas = self.temp_logger.get_temperaturas()

            # Log data at specified times
            if tempo_atual in self.tempos_registro and tempo_atual not in tempos_registrados:
                self.logger.debug(f"Logging data at time {tempo_atual}: {self.temperaturas}")
                self.dados.append((tempo_atual, self.temperaturas.copy()))
                tempos_registrados.add(tempo_atual)

            # Small pause
            time.sleep(0.1)

    def parar(self):
        """Stops the measurement process."""
        self.interrompido = True
        self.temp_logger.parar()
        self.logger.info("Measurement interrupted by user.")
