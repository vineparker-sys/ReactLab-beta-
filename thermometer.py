# thermometer.py

import re
import logging

class Termometro:
    """Processa e mantém as temperaturas lidas dos termopares."""

    def __init__(self, fila_dados):
        """Inicializa o termômetro com a fila de dados compartilhada."""
        self.fila_dados = fila_dados
        self.temperaturas = {"T1": "OFF", "T2": "OFF", "T3": "OFF", "T4": "OFF"}

    def extrair_temperaturas(self, dados):
        """Extrai as temperaturas dos dados brutos lidos."""
        termopares_novos = {}

        matches = re.findall(r'(\d{10,14})', dados)

        for match in matches:
            if len(match) == 14:
                canal = match[:2]
                temperatura_bruta = match[-3:]
                try:
                    temperatura = int(temperatura_bruta[:-1]) + (int(temperatura_bruta[-1]) / 10.0)
                except ValueError:
                    logging.warning(f"Erro ao converter temperatura bruta '{temperatura_bruta}' para canal '{canal}'")
                    continue

                if canal == "41":
                    termopares_novos["T1"] = f"{temperatura:.1f} °C"
                elif canal == "42":
                    termopares_novos["T2"] = f"{temperatura:.1f} °C"
                elif canal == "43":
                    termopares_novos["T3"] = f"{temperatura:.1f} °C"
                elif canal == "44":
                    termopares_novos["T4"] = f"{temperatura:.1f} °C"

        # Atualiza as temperaturas
        for termopar in ["T1", "T2", "T3", "T4"]:
            if termopar in termopares_novos:
                self.temperaturas[termopar] = termopares_novos[termopar]

    def ler_temperaturas(self):
        """Lê e processa todos os dados disponíveis na fila."""
        while not self.fila_dados.empty():
            dados_brutos = self.fila_dados.get()
            if dados_brutos:
                self.extrair_temperaturas(dados_brutos)
        return self.temperaturas.copy()
