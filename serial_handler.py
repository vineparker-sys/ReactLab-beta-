# serial_handler.py

import serial
import logging

class ManipuladorPortaSerial:
    """Gerencia a conexão serial com o dispositivo termômetro."""

    def __init__(self, porta_com, taxa_baud=9600):
        """Inicializa o manipulador com a porta e taxa de baud especificadas."""
        self.porta_com = porta_com
        self.taxa_baud = taxa_baud
        self.ser = None
        self.logger = logging.getLogger("serial_handler")

    def abrir(self):
        """Abre a conexão serial com as configurações especificadas."""
        try:
            self.ser = serial.Serial(
                port=self.porta_com,
                baudrate=self.taxa_baud,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1  # Timeout para evitar bloqueios
            )
            self.logger.debug(f"Porta serial {self.porta_com} aberta com sucesso.")
            return True
        except Exception as e:
            self.logger.error(f"Erro ao abrir a porta serial {self.porta_com}: {e}")
            print(f"Erro ao abrir a porta serial: {e}")
            return False

    def fechar(self):
        """Fecha a conexão serial se estiver aberta."""
        if self.ser and self.ser.is_open:
            self.ser.close()
            self.logger.debug("Porta serial fechada.")
            print("Porta serial fechada.")

    def ler_dados(self):
        """Lê dados da porta serial."""
        try:
            if self.ser and self.ser.in_waiting > 0:
                data = self.ser.read(self.ser.in_waiting)  # Lê todos os bytes disponíveis
                if data:
                    self.logger.debug(f"Dado lido da serial: {data}")
                    return data
            return None
        except Exception as e:
            self.logger.error(f"Erro ao ler da porta serial: {e}")
            return None
