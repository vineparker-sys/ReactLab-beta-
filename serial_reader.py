# serial_reader.py

import threading
import time
import logging

class LeitorSerialThread(threading.Thread):
    """Thread que realiza a leitura contínua da porta serial."""

    def __init__(self, manipulador_serial, fila_dados):
        super().__init__()
        self.manipulador_serial = manipulador_serial
        self.fila_dados = fila_dados
        self.running = True

    def run(self):
        """Método executado ao iniciar a thread."""
        while self.running:
            dados_brutos = self.manipulador_serial.ler_dados()
            if dados_brutos:
                self.fila_dados.put(dados_brutos)
            else:
                # Pequena pausa para evitar uso excessivo de CPU
                time.sleep(0.05)  # Reduzido para 50 ms

    def parar(self):
        """Interrompe a execução da thread."""
        self.running = False
