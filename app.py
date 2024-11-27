# app.py

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import threading
import time
import logging
import random  # Para simulação de temperaturas
import pandas as pd  # Para exportar para Excel
from PIL import Image
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import os
import getpass
from logging.handlers import TimedRotatingFileHandler

# Configuração de Logging
def configurar_logging():
    """Configura o sistema de logging com diferentes handlers para arquivo e console."""
    # Obter o nome do usuário atual
    username = getpass.getuser()

    # Nome do arquivo de log inclui o nome do usuário
    log_filename = f"debug_{username}.log"

    # Criar o logger raiz
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)  # Captura todos os níveis de log

    # Formato de log
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

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

configurar_logging()
logger = logging.getLogger(__name__)

# Definição das Fontes
FONTE_PADRAO = ("Roboto", 14)
FONTE_TITULO = ("Roboto", 20, "bold")
FONTE_SUBTITULO = ("Roboto", 16)

def formatar_tempo(segundos):
    """Formata o tempo em segundos para MM:SS."""
    minutos = segundos // 60
    segundos = segundos % 60
    return f"{minutos:02d}:{segundos:02d}"

class TemperatureLogger:
    """
    Classe para logar temperaturas reais via porta serial.
    """
    def __init__(self, porta, canais):
        self.porta = porta
        self.canais = canais
        self.temperaturas = {tp: 'OFF' for tp in canais.values()}
        self.running = False
        self.thread = None

    def start(self):
        """Inicia a leitura de temperaturas."""
        self.running = True
        self.thread = threading.Thread(target=self.leitura_serial)
        self.thread.daemon = True
        self.thread.start()
        logger.info("TemperatureLogger iniciado.")

    def leitura_serial(self):
        """Simula a leitura de temperaturas da porta serial."""
        while self.running:
            # Simulação: atribuir valores fictícios
            for canal, tp in self.canais.items():
                # Simula a leitura: gera um valor aleatório entre 20 e 100
                self.temperaturas[tp] = round(random.uniform(20, 100), 2)
            time.sleep(1)  # Simula intervalo de leitura

    def obter_temperaturas(self):
        """Obtém as temperaturas atuais."""
        return self.temperaturas.copy()

    def parar(self):
        """Para a leitura de temperaturas."""
        self.running = False
        if self.thread:
            self.thread.join()
        logger.info("TemperatureLogger parado.")

class SimulatedTemperatureLogger:
    """
    Simula a leitura de temperaturas para os termopares ativos.
    Útil quando a opção de simulação está ativada.
    """
    def __init__(self, termopares_ativos):
        self.termopares_ativos = termopares_ativos
        self.temperaturas = {tp: 25.0 for tp in termopares_ativos}
        self.running = False
        self.thread = None

    def start(self):
        """Inicia a simulação."""
        self.running = True
        self.thread = threading.Thread(target=self.simular_temperaturas)
        self.thread.daemon = True  # Permite que a thread seja encerrada com a aplicação
        self.thread.start()
        logger.info("Simulação de temperaturas iniciada.")

    def simular_temperaturas(self):
        """Aumenta as temperaturas simuladas de forma contínua a cada segundo."""
        while self.running:
            for tp in self.termopares_ativos:
                # Aumenta a temperatura de forma contínua
                self.temperaturas[tp] += 0.5  # Aumento fixo de 0.5°C por segundo
                # Limita as temperaturas a um intervalo realista
                self.temperaturas[tp] = min(100.0, self.temperaturas[tp])
            time.sleep(1)

    def obter_temperaturas(self):
        """Retorna as temperaturas atuais simuladas."""
        return self.temperaturas.copy()

    def parar(self):
        """Para a simulação."""
        self.running = False
        if self.thread:
            self.thread.join()
        logger.info("Simulação de temperaturas parada.")

class Medicao:
    """
    Classe para gerenciar a medição de temperaturas.
    """
    def __init__(self, temperature_logger, codigos_amostras):
        self.temperature_logger = temperature_logger
        self.codigos_amostras = codigos_amostras

    def obter_temperaturas(self):
        """Obtém as temperaturas atuais."""
        return self.temperature_logger.obter_temperaturas()

class ExportadorDados:
    """
    Classe para exportar os dados da medição para um arquivo Excel.
    """
    def __init__(self, dados, codigos_amostras_vars, tipo_analise="comum"):
        self.dados = dados
        self.codigos_amostras_vars = codigos_amostras_vars
        self.tipo_analise = tipo_analise

    def exportar_para_excel(self, filepath):
        """
        Exporta os dados para um arquivo Excel.
        """
        try:
            # Criar um DataFrame a partir dos dados
            df = pd.DataFrame(self.dados, columns=["Tempo", "Temperaturas"])

            # Expandir as colunas de temperaturas
            temp_data = df["Temperaturas"].apply(pd.Series)
            df = pd.concat([df["Tempo"], temp_data], axis=1)

            # Adicionar os IDs dos termopares
            for tp, var in self.codigos_amostras_vars.items():
                df[f"ID {tp}"] = var.get()

            # Salvar no arquivo Excel
            df.to_excel(filepath, index=False)

            return filepath
        except Exception as e:
            logger.error(f"Erro ao exportar para Excel: {e}")
            raise e

class ManipuladorPortaSerial:
    """
    Classe para manipular a porta serial.
    """
    def __init__(self, porta):
        self.porta = porta
        self.serial_connection = None

    def abrir(self):
        """Abre a porta serial."""
        try:
            import serial  # Importar aqui para evitar erro se não for usado
            self.serial_connection = serial.Serial(self.porta, 9600, timeout=1)
            logger.info(f"Porta serial {self.porta} aberta com sucesso.")
            return True
        except Exception as e:
            logger.error(f"Erro ao abrir a porta serial {self.porta}: {e}")
            return False

    def fechar(self):
        """Fecha a porta serial."""
        if self.serial_connection and self.serial_connection.is_open:
            self.serial_connection.close()
            logger.info(f"Porta serial {self.porta} fechada.")

    def read_data(self):
        """Lê dados da porta serial."""
        if self.serial_connection and self.serial_connection.is_open:
            try:
                data = self.serial_connection.readline().decode('utf-8').strip()
                return data
            except Exception as e:
                logger.error(f"Erro ao ler da porta serial {self.porta}: {e}")
                return None
        else:
            logger.warning(f"Porta serial {self.porta} não está aberta.")
            return None

class ReactLabApp(ctk.CTk):
    """
    Classe principal da aplicação ReactLab.
    Gerencia a interface gráfica, interações do usuário e integração com a lógica de medição.
    """
    def __init__(self):
        super().__init__()
        self.title("ReactLab - Analisador de Reatividade")
        
        # Carregar Tema Azul Claro Personalizado antes de configurar a janela
        self.carregar_tema_personalizado("blue-light")
        
        # Configurar a janela maximizada
        try:
            self.state("zoomed")  # Para Windows
        except tk.TclError:
            logger.warning("Não foi possível maximizar a janela usando 'zoomed'. Tentando com 'fullscreen'.")
            self.attributes('-fullscreen', True)
        
        # Atualizar para garantir que winfo_screenwidth e winfo_screenheight retornem valores corretos
        self.update_idletasks()
        
        # Variáveis de Controle
        self.planta_selecionada = ctk.StringVar(value="São José da Lapa")
        self.intervalo_selecionado = ctk.StringVar(value="30 segundos")
        self.analise_duracao_selected = ctk.StringVar(value="30 segundos")  # Variável para duração da análise
        self.simulacao_ativa = False  # Flag para simulação

        # Termopares Ativos
        self.termopares_ativos = []
        self.codigos_amostras_vars = {"T1": tk.StringVar(), "T2": tk.StringVar(), "T3": tk.StringVar(), "T4": tk.StringVar()}
        self.entry_widgets = {"T1": None, "T2": None, "T3": None, "T4": None}  # Para armazenar widgets de entrada

        # Dados de Medição
        self.dados = []  # Lista para armazenar dados de medição
        self.delta_t = 0.0  # Delta T calculado após a medição
        self.tempo_total = 0  # Tempo total da análise

        # Estado da Análise
        self.analise_em_andamento = False

        # Inicializar Interface
        self.create_menu()
        self.create_sidebar_left()
        self.create_sidebar_right()
        self.create_main_frame()
        self.create_bottom_frame()

        # Carregar Logos
        self.load_logos()

        # Protocolo para fechar a aplicação
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def carregar_tema_personalizado(self, tema_nome):
        """
        Carrega um tema personalizado a partir de um arquivo JSON.
        """
        try:
            tema_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "themes", f"{tema_nome}.json")
            if os.path.exists(tema_path):
                ctk.set_default_color_theme(tema_path)
                logger.info(f"Tema '{tema_nome}' carregado com sucesso.")
            else:
                logger.warning(f"Arquivo de tema '{tema_nome}.json' não encontrado. Usando tema padrão.")
        except Exception as e:
            logger.error(f"Erro ao carregar tema personalizado: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar tema personalizado:\n{e}")

    def create_menu(self):
        """
        Cria a barra de menus com a opção de Debug para ativar/desativar a simulação.
        """
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        # Menu Debug
        debug_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Debug", menu=debug_menu)

        # Adicionar checkbutton ao menu
        debug_menu.add_checkbutton(label="Ativar Simulação", command=self.toggle_simulacao)

    def toggle_simulacao(self):
        """
        Alterna a opção de simulação de temperaturas.
        """
        self.simulacao_ativa = not self.simulacao_ativa
        estado = "ativada" if self.simulacao_ativa else "desativada"
        messagebox.showinfo("Simulação", f"Simulação de temperaturas {estado}.")
        logger.info(f"Simulação de temperaturas {estado}.")

    def load_logos(self):
        """Carrega e exibe os logos da aplicação e da empresa."""
        # Diretório base
        base_path = os.path.dirname(os.path.abspath(__file__))
        assets_path = os.path.join(base_path, "assets")

        # Área para Logo e Nome do Programa na Barra Lateral Esquerda (Já foi movido para o topo na create_sidebar_left)
        # A seguir, removemos o posicionamento anterior do logo da empresa na janela principal
        # e o adicionamos à barra lateral direita.

        # Logo da Empresa (Lhoist) - Centralizado na barra lateral direita
        lhoist_logo_path = os.path.join(assets_path, "lhoist_logo.png")
        if os.path.exists(lhoist_logo_path):
            try:
                lhoist_logo_image = Image.open(lhoist_logo_path)
                lhoist_logo_image = lhoist_logo_image.resize((150, 75), Image.LANCZOS if hasattr(Image, 'LANCZOS') else Image.ANTIALIAS)
                self.lhoist_logo = ctk.CTkImage(light_image=lhoist_logo_image, dark_image=lhoist_logo_image, size=(150, 75))
                self.lhoist_logo_label = ctk.CTkLabel(self.sidebar_right, image=self.lhoist_logo, text="")
                self.lhoist_logo_label.pack(pady=10)
            except Exception as e:
                logger.error(f"Erro ao carregar logo da empresa: {e}")
                self.lhoist_logo_label = ctk.CTkLabel(self.sidebar_right, text="[Logo da Empresa]")
                self.lhoist_logo_label.pack(pady=10)
        else:
            logger.warning("Logo da empresa não encontrado.")
            self.lhoist_logo_label = ctk.CTkLabel(self.sidebar_right, text="[Logo da Empresa]")
            self.lhoist_logo_label.pack(pady=10)

    def create_sidebar_left(self):
        """
        Cria a barra lateral esquerda com opções de análise, botões de controle e logo no topo.
        """
        self.sidebar_left = ctk.CTkFrame(self, width=250, corner_radius=6)
        self.sidebar_left.pack(side="left", fill="y", padx=10, pady=10)

        # Seção de Logo e Nome do Programa no topo da barra lateral esquerda
        logo_nome_frame = ctk.CTkFrame(self.sidebar_left, corner_radius=6)
        logo_nome_frame.pack(pady=10, padx=10, fill="x")

        # Logo da Aplicação (ReactLab)
        base_path = os.path.dirname(os.path.abspath(__file__))
        assets_path = os.path.join(base_path, "assets")
        logo_path = os.path.join(assets_path, "logo.png")
        if os.path.exists(logo_path):
            try:
                logo_image = Image.open(logo_path)
                logo_image = logo_image.resize((100, 100), Image.LANCZOS if hasattr(Image, 'LANCZOS') else Image.ANTIALIAS)
                self.logo = ctk.CTkImage(light_image=logo_image, dark_image=logo_image, size=(100, 100))
                self.logo_label = ctk.CTkLabel(logo_nome_frame, image=self.logo, text="")
                self.logo_label.pack(pady=(0, 5))
            except Exception as e:
                logger.error(f"Erro ao carregar logo da aplicação: {e}")
                self.logo_label = ctk.CTkLabel(logo_nome_frame, text="[Logo da Aplicação]")
                self.logo_label.pack(pady=(0, 5))
        else:
            logger.warning("Logo da aplicação não encontrado.")
            self.logo_label = ctk.CTkLabel(logo_nome_frame, text="[Logo da Aplicação]")
            self.logo_label.pack(pady=(0, 5))

        # Nome do Programa
        self.nome_programa_label = ctk.CTkLabel(logo_nome_frame, text="ReactLab - Análise de Reatividade", font=("Roboto", 16, "bold"))
        self.nome_programa_label.pack()

        # Seção de Análise
        analise_frame = ctk.CTkFrame(self.sidebar_left, corner_radius=6)
        analise_frame.pack(pady=10, padx=10, fill="x")

        # Rótulo para Análise
        rotulo_analise = ctk.CTkLabel(analise_frame, text="Opções de Análise", font=FONTE_SUBTITULO)
        rotulo_analise.pack(pady=(0, 5), padx=10, anchor="w")

        # Botão Check para Termopares
        btn_check = ctk.CTkButton(analise_frame, text="Check Termopares", command=self.verificar_termopares, font=FONTE_PADRAO)
        btn_check.pack(pady=5, padx=10, fill="x")
        self.btn_check = btn_check

        # Espaço para Delta T após análise
        self.delta_label_sidebar = ctk.CTkLabel(self.sidebar_left, text="Delta T: N/A", font=FONTE_SUBTITULO)
        self.delta_label_sidebar.pack(pady=10, padx=20)

        # Linha de separação visual
        separator = ctk.CTkFrame(self.sidebar_left, height=2, fg_color="#D1D9E0")
        separator.pack(pady=10, padx=20, fill="x")

        # Botões de Controle
        controle_frame = ctk.CTkFrame(self.sidebar_left, corner_radius=6)
        controle_frame.pack(pady=10, padx=10, fill="x")

        # Botão para Iniciar Análise (sempre habilitado)
        btn_iniciar = ctk.CTkButton(controle_frame, text="Iniciar Análise", command=self.iniciar_analise, font=FONTE_PADRAO)
        btn_iniciar.pack(pady=5, padx=10, fill="x")
        self.btn_iniciar = btn_iniciar

        # Botão para Resetar Campos
        btn_resetar = ctk.CTkButton(controle_frame, text="Resetar Campos", command=self.resetar_campos, font=FONTE_PADRAO)
        btn_resetar.pack(pady=5, padx=10, fill="x")
        self.btn_resetar = btn_resetar

        # Botão para Exportar para Excel (desabilitado inicialmente)
        btn_exportar = ctk.CTkButton(controle_frame, text="Exportar para Excel", command=self.exportar_dados, state="disabled", font=FONTE_PADRAO)
        btn_exportar.pack(pady=5, padx=10, fill="x")
        self.btn_exportar = btn_exportar

    def create_sidebar_right(self):
        """
        Cria a barra lateral direita com grupos de seleção de planta, intervalo e duração da análise.
        Além disso, adiciona os campos de ID das amostras abaixo da seleção de duração.
        """
        self.sidebar_right = ctk.CTkFrame(self, width=250, corner_radius=6)
        self.sidebar_right.pack(side="right", fill="y", padx=10, pady=10)

        painel_direito = ctk.CTkFrame(self.sidebar_right, corner_radius=6)
        painel_direito.pack(side="top", fill="both", expand=True)

        # Grupo de Seleção de Planta
        grupo_planta = ctk.CTkFrame(painel_direito, corner_radius=6)
        grupo_planta.pack(pady=10, padx=10, fill="x")

        # Rótulo para o grupo de plantas
        rotulo_planta = ctk.CTkLabel(grupo_planta, text="Seleção de Planta", font=FONTE_SUBTITULO)
        rotulo_planta.pack(pady=(0, 5), padx=10, anchor="w")

        plantas = ["São José da Lapa", "Matozinhos", "Vitória"]
        for planta in plantas:
            radio = ctk.CTkRadioButton(grupo_planta, text=planta, variable=self.planta_selecionada, value=planta, font=FONTE_PADRAO, command=self.verificar_selecao)
            radio.pack(anchor="w", pady=2, padx=10)

        # Grupo de Seleção de Intervalo de Registro
        grupo_intervalo = ctk.CTkFrame(painel_direito, corner_radius=6)
        grupo_intervalo.pack(pady=10, padx=10, fill="x")

        # Rótulo para o grupo de intervalos
        rotulo_intervalo = ctk.CTkLabel(grupo_intervalo, text="Intervalo de Registro", font=FONTE_SUBTITULO)
        rotulo_intervalo.pack(pady=(0, 5), padx=10, anchor="w")

        intervalos = [
            ("30 segundos", "30 segundos"),
            ("1 minuto", "1 minuto")
        ]
        for texto, valor in intervalos:
            radio = ctk.CTkRadioButton(grupo_intervalo, text=texto, variable=self.intervalo_selecionado, value=valor, font=FONTE_PADRAO, command=self.verificar_selecao)
            radio.pack(anchor="w", pady=2, padx=10)

        # Grupo de Seleção de Duração da Análise
        grupo_duracao = ctk.CTkFrame(painel_direito, corner_radius=6)
        grupo_duracao.pack(pady=10, padx=10, fill="x")

        # Rótulo para o grupo de duração da análise
        rotulo_duracao = ctk.CTkLabel(grupo_duracao, text="Duração da Análise", font=FONTE_SUBTITULO)
        rotulo_duracao.pack(pady=(0, 5), padx=10, anchor="w")

        duracoes = [
            ("30 segundos", "30 segundos"),
            ("2 minutos", "2 minutos"),
            ("3 minutos", "3 minutos"),
            ("5 minutos", "5 minutos"),
            ("10 minutos", "10 minutos"),
            ("20 minutos", "20 minutos"),
            ("30 minutos", "30 minutos"),
            ("35 minutos", "35 minutos")
        ]
        for texto, valor in duracoes:
            radio = ctk.CTkRadioButton(grupo_duracao, text=texto, variable=self.analise_duracao_selected, value=valor, font=FONTE_PADRAO, command=self.verificar_selecao)
            radio.pack(anchor="w", pady=2, padx=10)

        # Campos de Preenchimento dos IDs das Amostras
        campos_id_frame = ctk.CTkFrame(painel_direito, corner_radius=6)
        campos_id_frame.pack(pady=20, padx=10, fill="x")

        rotulo_id = ctk.CTkLabel(campos_id_frame, text="IDs das Amostras", font=FONTE_SUBTITULO)
        rotulo_id.pack(pady=(0, 5), padx=10, anchor="w")

        for termopar in ["T1", "T2", "T3", "T4"]:
            frame = ctk.CTkFrame(campos_id_frame, corner_radius=6)
            frame.pack(pady=2, padx=10, fill="x")
            
            label = ctk.CTkLabel(frame, text=f"ID {termopar}:", font=FONTE_PADRAO)
            label.pack(side="left", padx=5, pady=5)

            entry = ctk.CTkEntry(frame, textvariable=self.codigos_amostras_vars[termopar], font=FONTE_PADRAO)
            entry.pack(side="left", padx=5, pady=5, fill="x", expand=True)
            entry.configure(state="disabled")  # Inicialmente desabilitado
            self.entry_widgets[termopar] = entry  # Armazenar widget de entrada

    def create_main_frame(self):
        """
        Cria o frame principal dividido em área de visualização.
        """
        main_frame = ctk.CTkFrame(self, corner_radius=6)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Área de Visualização de Resultados
        area_visualizacao = ctk.CTkFrame(main_frame, corner_radius=6)
        area_visualizacao.pack(pady=10, padx=10, fill="both", expand=True)

        # Frames para cada painel de visualização
        self.painel_grafico = ctk.CTkFrame(area_visualizacao, corner_radius=6)
        self.painel_tabela = ctk.CTkFrame(area_visualizacao, corner_radius=6)
        
        self.painel_grafico.pack(pady=10, padx=10, fill="both", expand=True)
        self.painel_tabela.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Inicializar Painéis
        self.create_painel_grafico()
        self.create_painel_tabela()

        # Mostrar ambos os painéis simultaneamente
        self.painel_grafico.tkraise()
        self.painel_tabela.tkraise()

    def create_painel_grafico(self):
        """
        Cria o painel do gráfico com temperaturas em tempo real.
        Inclui um gráfico Matplotlib e exibição das temperaturas de cada termopar.
        """
        # Gráfico Tempo x Temperatura
        self.fig, self.ax = plt.subplots(figsize=(8, 6))
        self.ax.set_title("Curva de Reatividade da Amostra", fontsize=16)
        self.ax.set_xlabel("Tempo (s)", fontsize=12)
        self.ax.set_ylabel("Temperatura (°C)", fontsize=12)
        self.ax.set_xlim(0, 300)  # Ajustar dinamicamente
        self.ax.set_ylim(20, 100)  # Ajustar conforme necessário

        # Definir cores para as linhas do gráfico
        cores = ["#e6194b", "#3cb44b", "#ffe119", "#4363d8"]  # Exemplo de cores

        self.lines = {}
        for i, termopar in enumerate(["T1", "T2", "T3", "T4"]):
            line, = self.ax.plot([], [], label=f"{termopar} (°C)", color=cores[i % len(cores)], linewidth=2)
            self.lines[termopar] = line
        self.ax.legend(fontsize=10)

        # Canvas do Gráfico
        self.canvas_grafico = FigureCanvasTkAgg(self.fig, master=self.painel_grafico)
        self.canvas_grafico.draw()
        self.canvas_grafico.get_tk_widget().pack(pady=10, padx=10, fill="both", expand=True)

        # Espaço para mostrar temperaturas em tempo real (Alterado para Grid)
        self.temp_labels = {}
        temp_frame = ctk.CTkFrame(self.painel_grafico)
        temp_frame.pack(pady=20, padx=10, anchor="center", fill="x")

        # Organizar as labels em grid (2 colunas)
        for index, termopar in enumerate(["T1", "T2", "T3", "T4"]):
            row = index // 2
            col = index % 2
            label = ctk.CTkLabel(temp_frame, text=f"{termopar}: 0.00°C", font=("Roboto", 16, "bold"))
            label.grid(row=row, column=col, padx=20, pady=10, sticky="nsew")
            self.temp_labels[termopar] = label

        # Configurar a expansão das colunas para que as labels ocupem o espaço disponível
        temp_frame.columnconfigure(0, weight=1)
        temp_frame.columnconfigure(1, weight=1)

    def create_painel_tabela(self):
        """
        Cria o painel da tabela que será preenchida conforme a análise avança.
        Utiliza um ttk.Treeview para exibição dos dados.
        """
        # Tabela de Dados usando ttk.Treeview
        self.table = ttk.Treeview(self.painel_tabela, columns=("Tempo", "T1", "T2", "T3", "T4"), show='headings')
        self.table.heading("Tempo", text="Tempo")
        self.table.heading("T1", text="T1 (°C)")
        self.table.heading("T2", text="T2 (°C)")
        self.table.heading("T3", text="T3 (°C)")
        self.table.heading("T4", text="T4 (°C)")

        # Definir a largura das colunas
        self.table.column("Tempo", width=100, anchor="center")
        self.table.column("T1", width=100, anchor="center")
        self.table.column("T2", width=100, anchor="center")
        self.table.column("T3", width=100, anchor="center")
        self.table.column("T4", width=100, anchor="center")

        # Adicionar uma barra de rolagem
        scrollbar = ttk.Scrollbar(self.painel_tabela, orient="vertical", command=self.table.yview)
        self.table.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.table.pack(pady=10, padx=10, fill="both", expand=True)

    def create_bottom_frame(self):
        """
        Cria a área inferior com barra de progresso, status e botão para interromper análise.
        Inicialmente, a barra de progresso e o botão são ocultos.
        """
        self.bottom_frame = ctk.CTkFrame(self, corner_radius=6)
        self.bottom_frame.pack(side="bottom", fill="x", padx=10, pady=10)

        # Barra de Progresso (Inicialmente Oculta)
        self.progress_bar = ctk.CTkProgressBar(self.bottom_frame, width=400, height=20)
        self.progress_bar.set(0)
        # Não fazer pack aqui

        # Rótulo de Status
        self.status_label = ctk.CTkLabel(self.bottom_frame, text="Status: Aguardando", font=FONTE_PADRAO)
        self.status_label.pack(side="left", padx=20)

        # Botão para Interromper Análise (Inicialmente Oculto)
        self.btn_interromper = ctk.CTkButton(
            self.bottom_frame,
            text="Interromper Análise",
            command=self.interromper_analise,
            fg_color="#FF4D4D",
            hover_color="#FF6666",
            font=FONTE_PADRAO
        )
        # Não fazer pack aqui

    def verificar_selecao(self, *args):
        """
        Verifica se todas as opções de RadioButtons estão selecionadas para habilitar o botão iniciar.
        """
        # Como o botão "Iniciar Análise" está sempre habilitado, não precisamos alterar seu estado aqui.
        pass

    def verificar_termopares(self):
        """
        Verifica os termopares ativos e habilita as caixas de texto correspondentes.
        """
        if self.simulacao_ativa:
            # Simulação: selecionar aleatoriamente termopares ativos
            self.termopares_ativos = [tp for tp in ["T1", "T2", "T3", "T4"] if random.choice([True, False])]
            logger.info(f"Simulação: Termopares ativos detectados: {self.termopares_ativos}")
        else:
            # Modo real: identificar termopares ativos via porta serial
            manipulador_serial = ManipuladorPortaSerial("COM12")  # Substitua "COM12" pela porta correta
            if not manipulador_serial.abrir():
                messagebox.showerror("Erro", "Não foi possível abrir a porta serial.")
                return

            canal_para_termopar = {"41": "T1", "42": "T2", "43": "T3", "44": "T4"}
            termopares_ativos = identificar_termopares_ativos(manipulador_serial, canal_para_termopar)
            manipulador_serial.fechar()

            if not termopares_ativos:
                messagebox.showwarning("Atenção", "Nenhum termopar ativo detectado.")
                self.termopares_ativos = []
            else:
                self.termopares_ativos = termopares_ativos
                logger.info(f"Termopares ativos detectados: {self.termopares_ativos}")

        # Habilitar ou desabilitar as caixas de texto com base nos termopares ativos
        for termopar, var in self.codigos_amostras_vars.items():
            entry_widget = self.entry_widgets.get(termopar)
            if termopar in self.termopares_ativos:
                # Habilitar entrada
                if entry_widget:
                    entry_widget.configure(state="normal")
            else:
                # Desabilitar e limpar entrada
                if entry_widget:
                    entry_widget.configure(state="disabled")
                    var.set("")  # Limpar o campo

    def iniciar_analise(self):
        """
        Inicia a análise após verificar as seleções e confirma com o usuário.
        """
        if not self.analise_em_andamento:
            # Verificar se todos os IDs dos termopares ativos estão preenchidos
            faltando_ids = [tp for tp in self.termopares_ativos if not self.codigos_amostras_vars[tp].get().strip()]
            if faltando_ids:
                messagebox.showwarning("Campos Incompletos", f"Por favor, preencha os IDs para os termopares ativos: {', '.join(faltando_ids)}.")
                return

            # Confirmar com o usuário antes de iniciar
            nomes_amostras = ", ".join([f"ID {tp}: {var.get().strip()}" for tp, var in self.codigos_amostras_vars.items() if tp in self.termopares_ativos])
            msg = (f"Você está prestes a iniciar a análise de reatividade para as amostras:\n{nomes_amostras}.\n\n"
                   "Para iniciar a análise, clique em OK.\n"
                   "Para cancelar, clique em Cancelar.")
            resposta = messagebox.askokcancel("Confirmar Início da Medição", msg)
            if resposta:
                logger.info("Usuário confirmou o início da medição.")
                self.analise_em_andamento = True
                self.btn_iniciar.configure(state="disabled")
                self.btn_exportar.configure(state="disabled")
                self.status_label.configure(text="Status: Analisando...")

                # Exibir a barra de progresso e o botão de interrupção
                self.progress_bar.pack(side="left", padx=20, pady=10, fill="x", expand=True)
                self.btn_interromper.pack(side="right", padx=20)

                self.iniciar_mediacao()
            else:
                logger.info("Usuário cancelou o início da medição.")

    def iniciar_mediacao(self):
        """
        Inicia a medição em uma thread separada.
        """
        # Resetar dados anteriores
        self.dados = []
        # Resetar gráfico e tabela antes de iniciar
        self.resetar_grafico()
        self.resetar_tabela()

        # Coletar IDs das amostras
        codigos_amostras = {tp: var.get().strip() for tp, var in self.codigos_amostras_vars.items() if tp in self.termopares_ativos}

        # Inicializar TemperatureLogger ou SimulatedTemperatureLogger
        if self.simulacao_ativa:
            self.temp_logger_simulado = SimulatedTemperatureLogger(self.termopares_ativos)
            self.temp_logger_simulado.start()
            temperature_logger_instance = self.temp_logger_simulado
        else:
            # Usar o TemperatureLogger real
            manipulador_serial = ManipuladorPortaSerial("COM12")  # Substitua "COM12" pela porta correta
            if not manipulador_serial.abrir():
                messagebox.showerror("Erro", "Não foi possível abrir a porta serial.")
                self.analise_em_andamento = False
                self.btn_iniciar.configure(state="normal")
                self.btn_exportar.configure(state="disabled")
                self.status_label.configure(text="Status: Aguardando")
                return
            canal_para_termopar = {"41": "T1", "42": "T2", "43": "T3", "44": "T4"}
            temperature_logger_instance = TemperatureLogger(manipulador_serial.porta, canal_para_termopar)
            temperature_logger_instance.start()

        # Inicializar Medicao
        self.medicao = Medicao(temperature_logger_instance, codigos_amostras)

        # Definir duração da medição com base na seleção
        duracao = self.analise_duracao_selected.get()
        if duracao == "30 segundos":
            tempo_total = 30
        elif duracao == "2 minutos":
            tempo_total = 120
        elif duracao == "3 minutos":
            tempo_total = 180
        elif duracao == "5 minutos":
            tempo_total = 300
        elif duracao == "10 minutos":
            tempo_total = 600
        elif duracao == "20 minutos":
            tempo_total = 1200
        elif duracao == "30 minutos":
            tempo_total = 1800
        elif duracao == "35 minutos":
            tempo_total = 2100
        else:
            # Caso não haja correspondência, definir valores padrão
            tempo_total = 30

        self.tempo_total = tempo_total  # Armazenar o tempo total para uso na barra de progresso

        # Obter o intervalo de registro
        intervalo = self.intervalo_selecionado.get()
        if intervalo == "30 segundos":
            intervalo_segundos = 30
        elif intervalo == "1 minuto":
            intervalo_segundos = 60
        else:
            intervalo_segundos = 30  # Valor padrão

        # Calcular todos os tempos de registro
        log_times = []
        current_time = 0
        while current_time <= tempo_total:
            log_times.append(current_time)
            current_time += intervalo_segundos

        # Garantir que o tempo_total esteja incluído
        if log_times[-1] != tempo_total:
            log_times.append(tempo_total)

        # Ordenar e remover duplicatas
        log_times = sorted(list(set(log_times)))

        # Iniciar a thread de medição
        self.thread_medicao = threading.Thread(target=self.run_medicao, args=(tempo_total, log_times))
        self.thread_medicao.daemon = True  # Permite que a thread seja encerrada com a aplicação
        self.thread_medicao.start()

    def run_medicao(self, tempo_total, log_times):
        """
        Executa a medição e atualiza a interface conforme os dados são coletados.
        """
        try:
            start_time = time.time()
            for log_time in log_times:
                if not self.analise_em_andamento:
                    break  # Interromper a medição se analise_em_andamento for False

                # Esperar até o próximo tempo de registro
                time_to_wait = log_time - (time.time() - start_time)
                if time_to_wait > 0:
                    time.sleep(time_to_wait)

                elapsed_time = int(time.time() - start_time)

                # Registrar a temperatura
                temperaturas = self.medicao.obter_temperaturas()
                self.dados.append((elapsed_time, temperaturas))
                # Atualizar a interface na thread principal
                self.after(0, self.atualizar_gui, elapsed_time, temperaturas)
                logger.info(f"Registro de temperatura em {formatar_tempo(elapsed_time)}: {temperaturas}")

            # Finalizar a medição
            if self.analise_em_andamento:
                self.delta_t = self.calcular_delta_t()
                self.after(0, self.atualizar_delta_t_sidebar, self.delta_t)
                self.after(0, self.btn_exportar.configure, {'state': "normal"})
                self.after(0, self.status_label.configure, {'text': "Status: Concluído"})
                messagebox.showinfo("Medição", "Medição concluída com sucesso!")
                logger.info("Medição concluída com sucesso.")
        except Exception as e:
            logger.error(f"Erro durante a medição: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro durante a medição:\n{e}")
            self.after(0, self.status_label.configure, {'text': "Status: Erro na Medição"})
        finally:
            self.analise_em_andamento = False
            # Parar o TemperatureLogger real se não estiver em simulação
            if not self.simulacao_ativa and hasattr(self, 'medicao') and hasattr(self.medicao, 'temperature_logger'):
                self.medicao.temperature_logger.parar()
            # Ocultar a barra de progresso e o botão de interrupção
            self.after(0, self.progress_bar.pack_forget)
            self.after(0, self.btn_interromper.pack_forget)

    def atualizar_delta_t_sidebar(self, delta_t):
        """
        Atualiza o rótulo de Delta T na sidebar esquerda.
        """
        self.delta_label_sidebar.configure(text=f"Delta T: {delta_t:.2f} °C")

    def calcular_delta_t(self):
        """
        Calcula o Delta T com base nos dados coletados.
        """
        if self.dados:
            temperaturas_iniciais = [temp for temp in self.dados[0][1].values()]
            temperaturas_finais = [temp for temp in self.dados[-1][1].values()]
            delta = sum([f - i for i, f in zip(temperaturas_iniciais, temperaturas_finais)]) / len(temperaturas_iniciais)
            return delta
        return 0.0

    def exportar_dados(self):
        """
        Exporta os dados da medição para um arquivo Excel.
        """
        try:
            if not self.dados:
                messagebox.showwarning("Exportar", "Não há dados para exportar.")
                return

            # Definir onde salvar o arquivo
            filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                    filetypes=[("Excel files", "*.xlsx")],
                                                    initialfile="Resultados_Medicao.xlsx")
            if not filepath:
                return  # O usuário cancelou a operação

            # Inicializar ExportadorDados
            exportador = ExportadorDados(self.dados, self.codigos_amostras_vars, tipo_analise="comum")
            caminho_exportacao = exportador.exportar_para_excel(filepath)

            messagebox.showinfo("Exportar", f"Dados exportados com sucesso para {caminho_exportacao}")
            logger.info(f"Dados exportados para {caminho_exportacao}")
        except Exception as e:
            logger.error(f"Erro ao exportar para Excel: {e}", exc_info=True)
            messagebox.showerror("Erro", f"Ocorreu um erro ao exportar os dados:\n{e}")

    def resetar_campos(self):
        """
        Reseta todas as seleções, entradas e estados da interface.
        """
        self.planta_selecionada.set("São José da Lapa")
        self.intervalo_selecionado.set("30 segundos")
        self.analise_duracao_selected.set("30 segundos")
        self.btn_iniciar.configure(state="normal")
        self.btn_exportar.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="Status: Aguardando")
        self.delta_label_sidebar.configure(text="Delta T: N/A")

        # Ocultar a barra de progresso e o botão de interrupção se estiverem visíveis
        self.progress_bar.pack_forget()
        self.btn_interromper.pack_forget()

        # Resetar entradas de IDs
        for termopar, var in self.codigos_amostras_vars.items():
            entry_widget = self.entry_widgets.get(termopar)
            if entry_widget:
                entry_widget.configure(state="disabled")
                var.set("")

        # Resetar dados
        self.dados = []

        # Resetar gráficos e tabela
        self.resetar_grafico()
        self.resetar_tabela()

        # Resetar temperaturas
        for termopar, label in self.temp_labels.items():
            label.configure(text=f"{termopar}: 0.00°C")

        logger.info("Campos resetados pelo usuário.")

    def mudar_tema(self, tema):
        """
        Muda o tema da aplicação.
        """
        # Removido: OptionMenu para seleção de tema
        pass

    def interromper_analise(self):
        """
        Interrompe a análise em andamento.
        """
        if self.analise_em_andamento:
            resposta = messagebox.askyesno("Interromper Análise", "Tem certeza que deseja interromper a análise?")
            if resposta:
                self.analise_em_andamento = False
                if self.simulacao_ativa and hasattr(self, 'temp_logger_simulado'):
                    self.temp_logger_simulado.parar()
                elif hasattr(self, 'medicao') and hasattr(self.medicao, 'temperature_logger'):
                    self.medicao.temperature_logger.parar()
                logger.info("Análise interrompida pelo usuário.")
                messagebox.showinfo("Interromper Análise", "Análise interrompida pelo usuário.")
                # Resetar barra de progresso e status
                self.progress_bar.set(0)
                self.status_label.configure(text="Status: Análise Interrompida")
                # Ocultar a barra de progresso e o botão de interrupção
                self.progress_bar.pack_forget()
                self.btn_interromper.pack_forget()

    def atualizar_gui(self, tempo, temperaturas):
        """
        Atualiza o gráfico, tabela, barra de progresso e temperaturas em tempo real.
        """
        # Atualizar Gráfico
        tempos = [d[0] for d in self.dados]
        for termopar in self.termopares_ativos:
            self.lines[termopar].set_data(tempos, [d[1].get(termopar, 0) for d in self.dados])
        self.ax.set_xlim(0, max(self.tempo_total, max(tempos) + 10))
        self.ax.relim()
        self.ax.autoscale_view()
        self.canvas_grafico.draw()

        # Atualizar Temperaturas em Tempo Real
        for termopar in self.termopares_ativos:
            temp = temperaturas.get(termopar, 0.0)
            self.temp_labels[termopar].configure(text=f"{termopar}: {temp:.2f}°C")

        # Atualizar Barra de Progresso
        percent = (tempo / self.tempo_total) * 100 if self.tempo_total > 0 else 0
        self.progress_bar.set(percent / 100)

        # Atualizar Tabela
        tempo_str = formatar_tempo(tempo)
        temp_values = [f"{temperaturas.get(tp, 0):.2f}" for tp in ["T1", "T2", "T3", "T4"]]
        self.table.insert("", "end", values=(tempo_str, *temp_values))

    def resetar_grafico(self):
        """
        Reseta o gráfico para um estado vazio.
        """
        self.ax.cla()
        self.ax.set_title("Curva de Reatividade da Amostra", fontsize=16)
        self.ax.set_xlabel("Tempo (s)", fontsize=12)
        self.ax.set_ylabel("Temperatura (°C)", fontsize=12)
        self.ax.set_xlim(0, 300)
        self.ax.set_ylim(20, 100)
        self.lines = {}
        cores = ["#e6194b", "#3cb44b", "#ffe119", "#4363d8"]
        for i, termopar in enumerate(["T1", "T2", "T3", "T4"]):
            line, = self.ax.plot([], [], label=f"{termopar} (°C)", color=cores[i % len(cores)], linewidth=2)
            self.lines[termopar] = line
        self.ax.legend(fontsize=10)
        self.canvas_grafico.draw()

    def resetar_tabela(self):
        """
        Reseta a tabela para um estado vazio.
        """
        for item in self.table.get_children():
            self.table.delete(item)

    def on_closing(self):
        """
        Função chamada quando a janela está sendo fechada.
        Garante que todas as threads sejam encerradas corretamente.
        """
        if self.analise_em_andamento:
            resposta = messagebox.askyesno("Sair", "Uma análise está em andamento. Deseja interromper e sair?")
            if resposta:
                self.analise_em_andamento = False
                if self.simulacao_ativa and hasattr(self, 'temp_logger_simulado'):
                    self.temp_logger_simulado.parar()
                elif hasattr(self, 'medicao') and hasattr(self.medicao, 'temperature_logger'):
                    self.medicao.temperature_logger.parar()
                self.destroy()
        else:
            self.destroy()

def identificar_termopares_ativos(manipulador_serial, canal_para_termopar):
    """Identifica os termopares ativos usando o TemperatureLogger por alguns segundos."""
    logger = logging.getLogger("main")
    logger.info("Iniciando identificação dos termopares ativos...")
    
    # Inicializar o TemperatureLogger com todos os termopares possíveis
    todos_termopares = list(canal_para_termopar.values())
    temp_logger = TemperatureLogger(manipulador_serial.porta, canal_para_termopar)
    temp_logger.start()
    time.sleep(2)  # Esperar por 2 segundos para coletar dados
    temp_logger.parar()

    # Verificar quais termopares têm valores diferentes de 'OFF'
    termopares_ativos = [tp for tp, temp in temp_logger.obter_temperaturas().items() if temp != 'OFF']
    logger.info(f"Termopares ativos detectados: {termopares_ativos}")
    return termopares_ativos

if __name__ == "__main__":
    # Definir o modo de aparência para 'light' antes de iniciar a aplicação
    ctk.set_appearance_mode("light")

    # Verificar se o tema personalizado existe
    tema_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "themes", "blue-light.json")
    if not os.path.exists(tema_path):
        messagebox.showerror("Erro", "Arquivo de tema 'blue-light.json' não encontrado na pasta 'themes'.")
    else:
        app = ReactLabApp()
        app.mainloop()
