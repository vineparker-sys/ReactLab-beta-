# app.py

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import threading
import time
import logging
import random  # For temperature simulation
from PIL import Image
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import os
import getpass
from logging.handlers import TimedRotatingFileHandler
import sys  # For sys.exit()

# Importing modules from the project
from measurement import Medicao
from data_exporter import ExportadorDados
from temperature_logger import TemperatureLogger, SimulatedTemperatureLogger
from serial_handler import ManipuladorPortaSerial
from utils import formatar_tempo
from main import identificar_termopares_ativos
from localization import Localizer
import win32com.client as win32  # For interacting with Outlook

# Definição das fontes padrão
FONTE_PADRAO = ("Roboto", 14)
FONTE_TITULO = ("Roboto", 20, "bold")
FONTE_SUBTITULO = ("Roboto", 16)

# Logging Configuration
logger = logging.getLogger(__name__)

def configurar_logging():
    """Configura o sistema de logging com diferentes handlers para arquivo e console."""
    import os
    import getpass

    # Obter o nome do usuário e pasta de dados local
    username = getpass.getuser()
    local_appdata = os.getenv('LOCALAPPDATA', os.path.expanduser('~\\AppData\\Local'))

    # Diretório para armazenar o log
    log_dir = os.path.join(local_appdata, "ReactLab")
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)  # Cria o diretório se não existir

    # Caminho completo para o arquivo de log
    log_filename = os.path.join(log_dir, f"debug_{username}.log")

    # Configurar o logger
    logger.setLevel(logging.DEBUG)  # Captura todos os níveis de log
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    file_handler = TimedRotatingFileHandler(log_filename, when="midnight", interval=1, backupCount=7)
    file_handler.setLevel(logging.DEBUG)  # Log todos os níveis no arquivo
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)  # Apenas WARNING ou acima no console
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

# Chamar a função para configurar o logging
configurar_logging()

class ReactLabApp(ctk.CTk):
    """
    Main class for the ReactLab application.
    Manages the GUI, user interactions, and integration with measurement logic.
    """
    def __init__(self):
        super().__init__()
        self.title("ReactLab - Analisador de Reatividade")
        
        # Load Custom Light Blue Theme before configuring the window
        self.carregar_tema_personalizado("blue-light")
        
        # Maximize the window
        try:
            self.state("zoomed")  # For Windows
        except tk.TclError:
            logger.warning("Cannot maximize window using 'zoomed'. Trying 'fullscreen'.")
            self.attributes('-fullscreen', True)
        
         # Definir o ícone da janela (barra de tarefas e topo da janela)
        base_path = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_path, "assets", "logo.ico")  
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)  # Para ícones .ico
        
        # Update to ensure winfo_screenwidth and winfo_screenheight return correct values
        self.update_idletasks()
        
        # Initialize Localizer
        self.localizer = Localizer(default_language="pt")
        self.language_selected = tk.StringVar(value="pt")  # Default language: Portuguese

        # Control Variables
        self.planta_selecionada = ctk.StringVar(value="São José da Lapa")
        self.intervalo_selecionado = ctk.StringVar(value="30 segundos")
        self.analise_duracao_selected = ctk.StringVar(value="30 segundos")  # Analysis duration variable
        self.simulacao_ativa = False  # Simulation flag

        # Active Thermocouples
        self.termopares_ativos = []
        self.codigos_amostras_vars = {"T1": tk.StringVar(), "T2": tk.StringVar(), "T3": tk.StringVar(), "T4": tk.StringVar()}
        self.entry_widgets = {"T1": None, "T2": None, "T3": None, "T4": None}  # To store entry widgets

        # Measurement Data
        self.dados = []  # List to store measurement data
        self.delta_t = 0.0  # Delta T calculated after measurement
        self.tempo_total = 0  # Total analysis time

        # Analysis State
        self.analise_em_andamento = False

        # Initialize Interface
        self.create_menu()
        self.create_sidebar_left()
        self.create_sidebar_right()
        self.create_main_frame()
        self.create_bottom_frame()

        # Load Logos
        self.load_logos()

        # Protocol to close the application
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def carregar_tema_personalizado(self, tema_nome):
        """
        Loads a custom theme from a JSON file.
        """
        try:
            tema_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "themes", f"{tema_nome}.json")
            if os.path.exists(tema_path):
                ctk.set_default_color_theme(tema_path)
                logger.info(f"Theme '{tema_nome}' loaded successfully.")
            else:
                logger.warning(f"Theme file '{tema_nome}.json' not found. Using default theme.")
        except Exception as e:
            logger.error(f"Error loading custom theme: {e}")
            messagebox.showerror(self.localizer.translate("error"), f"{self.localizer.translate('error_loading_theme')}\n{e}")

    def create_menu(self):
        """
        Creates the menu bar with File, Debug, Help, and Language menus.
        """
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)

        # File Menu
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.localizer.translate("file_menu"), menu=self.file_menu)
        self.file_menu.add_command(label=self.localizer.translate("reset_fields"), command=self.resetar_campos)
        self.file_menu.add_command(label=self.localizer.translate("export_data"), command=self.exportar_dados)
        self.file_menu.add_separator()
        self.file_menu.add_command(label=self.localizer.translate("exit"), command=self.on_closing)

        # Debug Menu
        self.debug_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.localizer.translate("debug_menu"), menu=self.debug_menu)
        self.debug_menu.add_checkbutton(label=self.localizer.translate("activate_simulation"), command=self.toggle_simulacao)

        # Help Menu
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.localizer.translate("help_menu"), menu=self.help_menu)
        self.help_menu.add_command(label=self.localizer.translate("about"), command=self.mostrar_sobre)
        self.help_menu.add_command(label=self.localizer.translate("support"), command=self.abrir_suporte)

        # Language Menu
        self.language_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=self.localizer.translate("language_menu"), menu=self.language_menu)
        for lang_code, lang_name in [("en", "English"), ("pt", "Português"), ("fr", "Français")]:
            self.language_menu.add_radiobutton(label=lang_name, variable=self.language_selected, value=lang_code,
                                            command=self.change_language)

    def change_language(self):
        """
        Changes the application's language.
        """
        selected_language = self.language_selected.get()
        try:
            self.localizer.set_language(selected_language)
            # Rebuild the menu bar
            self.rebuild_menu()
            self.update_texts()
            logger.info(f"Language changed to: {selected_language}")
        except ValueError as e:
            messagebox.showerror(self.localizer.translate("error"), str(e))
            logger.error(f"Error changing language: {e}")
            
    
    def rebuild_menu(self):
        """
        Rebuilds the menu bar with the new language settings.
        """
        # Remove the existing menu bar
        self.menu_bar.destroy()
        # Recreate the menu bar
        self.create_menu()


    def update_texts(self):
        """
        Updates all texts in the application based on the selected language.
        """
        # Update other widgets
        self.nome_programa_label.configure(text=self.localizer.translate("app_name"))
        self.btn_check.configure(text=self.localizer.translate("check_thermocouples"))
        self.btn_iniciar.configure(text=self.localizer.translate("start_analysis"))
        self.btn_resetar.configure(text=self.localizer.translate("reset_fields"))
        self.btn_exportar.configure(text=self.localizer.translate("export_data"))
        self.delta_label_sidebar.configure(text=f"{self.localizer.translate('delta_t')}: N/A")
        self.status_label.configure(text=f"{self.localizer.translate('status')}: {self.localizer.translate('waiting')}")
        self.btn_interromper.configure(text=self.localizer.translate("interrupt_analysis"))

        # Update labels in sidebars, main frame, etc.
        # Left Sidebar labels
        self.rotulo_analise.configure(text=self.localizer.translate("analysis_options"))

        # Right Sidebar labels
        self.rotulo_planta.configure(text=self.localizer.translate("plant_selection"))
        self.rotulo_intervalo.configure(text=self.localizer.translate("record_interval"))
        self.rotulo_duracao.configure(text=self.localizer.translate("analysis_duration"))
        self.rotulo_id.configure(text=self.localizer.translate("sample_ids"))

        for termopar, label in self.sample_id_labels.items():
            label.configure(text=f"{self.localizer.translate('id')} {termopar}:")

        # Update Radio Buttons
        self.update_radio_buttons()

        # Update real-time temperature labels
        for termopar, label in self.temp_labels.items():
            label.configure(text=f"{termopar}: 0.00°C")

        # Update graph labels
        self.update_graph_labels()

        # Update table headings
        self.update_table_headings()

    def update_radio_buttons(self):
        # Update plant radio buttons
        plant_texts = [self.localizer.translate("plant_sjl"), self.localizer.translate("plant_mtz"), self.localizer.translate("plant_vitoria")]
        for radio, text in zip(self.plant_radiobuttons, plant_texts):
            radio.configure(text=text)

        # Update interval radio buttons
        interval_texts = [self.localizer.translate("30_seconds"), self.localizer.translate("1_minute")]
        for radio, text in zip(self.interval_radiobuttons, interval_texts):
            radio.configure(text=text)

        # Update duration radio buttons
        duration_texts = [self.localizer.translate("30_seconds"), self.localizer.translate("2_minutes"), self.localizer.translate("3_minutes"),
                        self.localizer.translate("5_minutes"), self.localizer.translate("10_minutes"), self.localizer.translate("20_minutes"),
                        self.localizer.translate("30_minutes"), self.localizer.translate("35_minutes")]
        for radio, text in zip(self.duration_radiobuttons, duration_texts):
            radio.configure(text=text)

    def update_graph_labels(self):
        self.ax.set_title(self.localizer.translate("graph_title"), fontsize=16)
        self.ax.set_xlabel(self.localizer.translate("time_seconds"), fontsize=12)
        self.ax.set_ylabel(self.localizer.translate("temperature_celsius"), fontsize=12)
        self.ax.legend(fontsize=10)
        self.canvas_grafico.draw()

    def update_table_headings(self):
            self.table.heading("Tempo", text=self.localizer.translate("time"))
            self.table.heading("T1", text=f"T1 ({self.localizer.translate('temperature_celsius')})")
            self.table.heading("T2", text=f"T2 ({self.localizer.translate('temperature_celsius')})")
            self.table.heading("T3", text=f"T3 ({self.localizer.translate('temperature_celsius')})")
            self.table.heading("T4", text=f"T4 ({self.localizer.translate('temperature_celsius')})")

        
    def toggle_simulacao(self):
        """
        Toggles the temperature simulation option.
        """
        self.simulacao_ativa = not self.simulacao_ativa
        estado = self.localizer.translate("activated") if self.simulacao_ativa else self.localizer.translate("deactivated")
        messagebox.showinfo(self.localizer.translate("simulation"), f"{self.localizer.translate('simulation_temperatures')} {estado}.")
        logger.info(f"Simulação de temperaturas {estado}.")

    def mostrar_sobre(self):
        """
        Shows information about the program.
        """
        info_text = self.localizer.translate("about_text")
        messagebox.showinfo(self.localizer.translate("about"), info_text)

    def abrir_suporte(self):
        """
        Opens a new email in Outlook to send an email to support.
        """
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'suporte@exemplo.com'  # Substitua pelo e-mail de suporte
            mail.CC = 'gerencia@exemplo.com'  # Substitua pelo e-mail de gerência
            mail.Subject = self.localizer.translate('support_email_subject')
            mail.Body = self.localizer.translate('support_email_body')
            mail.Display()
            logger.info("Outlook opened for support email.")
        except Exception as e:
            messagebox.showerror(self.localizer.translate("error"), f"{self.localizer.translate('cannot_open_outlook')}\n{e}")
            logger.error(f"Error opening Outlook: {e}")

    def load_logos(self):
        """Loads and displays the application and company logos."""
        # Base directory
        base_path = os.path.dirname(os.path.abspath(__file__))
        assets_path = os.path.join(base_path, "assets")

        # Logo of the Company - Centered in the right sidebar
        company_logo_path = os.path.join(assets_path, "lhoist_logo.png") 
        if os.path.exists(company_logo_path):
            try:
                company_logo_image = Image.open(company_logo_path)
                company_logo_image = company_logo_image.resize((150, 75), Image.LANCZOS if hasattr(Image, 'LANCZOS') else Image.ANTIALIAS)
                self.company_logo = ctk.CTkImage(light_image=company_logo_image, dark_image=company_logo_image, size=(150, 75))
                self.company_logo_label = ctk.CTkLabel(self.sidebar_right, image=self.company_logo, text="")
                self.company_logo_label.pack(pady=10)
            except Exception as e:
                logger.error(f"Error loading company logo: {e}")
                self.company_logo_label = ctk.CTkLabel(self.sidebar_right, text="[Logo da Empresa]")
                self.company_logo_label.pack(pady=10)
        else:
            logger.warning("Company logo not found.")
            self.company_logo_label = ctk.CTkLabel(self.sidebar_right, text="[Logo da Empresa]")
            self.company_logo_label.pack(pady=10)

    def create_sidebar_left(self):
        """
        Creates the left sidebar with analysis options, control buttons, and logo at the top.
        """
        self.sidebar_left = ctk.CTkFrame(self, width=250, corner_radius=6)
        self.sidebar_left.pack(side="left", fill="y", padx=10, pady=10)

        # Logo and Program Name section at the top of the left sidebar
        logo_nome_frame = ctk.CTkFrame(self.sidebar_left, corner_radius=6)
        logo_nome_frame.pack(pady=10, padx=10, fill="x")

        # Application Logo
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
                logger.error(f"Error loading application logo: {e}")
                self.logo_label = ctk.CTkLabel(logo_nome_frame, text="[Logo da Aplicação]")
                self.logo_label.pack(pady=(0, 5))
        else:
            logger.warning("Application logo not found.")
            self.logo_label = ctk.CTkLabel(logo_nome_frame, text="[Logo da Aplicação]")
            self.logo_label.pack(pady=(0, 5))

        # Program Name
        self.nome_programa_label = ctk.CTkLabel(logo_nome_frame, text=self.localizer.translate("app_name"), font=("Roboto", 16, "bold"))
        self.nome_programa_label.pack()

       # Analysis Section
        self.analise_frame = ctk.CTkFrame(self.sidebar_left, corner_radius=6)
        self.analise_frame.pack(pady=10, padx=10, fill="x")

        # Label for Analysis
        self.rotulo_analise = ctk.CTkLabel(self.analise_frame, text=self.localizer.translate("analysis_options"), font=FONTE_SUBTITULO)
        self.rotulo_analise.pack(pady=(0, 5), padx=10, anchor="w")


        # Button Check for Thermocouples
        self.btn_check = ctk.CTkButton(self.analise_frame, text=self.localizer.translate("check_thermocouples"), command=self.verificar_termopares, font=FONTE_PADRAO)
        self.btn_check.pack(pady=5, padx=10, fill="x")

        # Space for Delta T after analysis
        self.delta_label_sidebar = ctk.CTkLabel(self.sidebar_left, text=f"{self.localizer.translate('delta_t')}: N/A", font=FONTE_SUBTITULO)
        self.delta_label_sidebar.pack(pady=10, padx=20)

        # Visual separation line
        separator = ctk.CTkFrame(self.sidebar_left, height=2, fg_color="#D1D9E0")
        separator.pack(pady=10, padx=20, fill="x")

        # Control Buttons
        controle_frame = ctk.CTkFrame(self.sidebar_left, corner_radius=6)
        controle_frame.pack(pady=10, padx=10, fill="x")

        # Button to Start Analysis (always enabled)
        self.btn_iniciar = ctk.CTkButton(controle_frame, text=self.localizer.translate("start_analysis"), command=self.iniciar_analise, font=FONTE_PADRAO)
        self.btn_iniciar.pack(pady=5, padx=10, fill="x")

        # Button to Reset Fields
        self.btn_resetar = ctk.CTkButton(controle_frame, text=self.localizer.translate("reset_fields"), command=self.resetar_campos, font=FONTE_PADRAO)
        self.btn_resetar.pack(pady=5, padx=10, fill="x")

        # Button to Export to Excel (initially disabled)
        self.btn_exportar = ctk.CTkButton(controle_frame, text=self.localizer.translate("export_data"), command=self.exportar_dados, state="disabled", font=FONTE_PADRAO)
        self.btn_exportar.pack(pady=5, padx=10, fill="x")

    def create_sidebar_right(self):
        """
        Creates the right sidebar with plant selection groups, interval, and analysis duration.
        Also adds the sample ID fields below the duration selection.
        """
        self.sidebar_right = ctk.CTkFrame(self, width=250, corner_radius=6)
        self.sidebar_right.pack(side="right", fill="y", padx=10, pady=10)

        painel_direito = ctk.CTkFrame(self.sidebar_right, corner_radius=6)
        painel_direito.pack(side="top", fill="both", expand=True)

        # Plant Selection Group Frame
        grupo_planta = ctk.CTkFrame(painel_direito, corner_radius=6)
        grupo_planta.pack(pady=10, padx=10, fill="x")

        # Label for plant group
        self.rotulo_planta = ctk.CTkLabel(grupo_planta, text=self.localizer.translate("plant_selection"), font=FONTE_SUBTITULO)
        self.rotulo_planta.pack(pady=(0, 5), padx=10, anchor="w")

        # Plant Selection Radio Buttons
        self.plant_radiobuttons = []
        plantas = [self.localizer.translate("plant_sjl"), self.localizer.translate("plant_mtz"), self.localizer.translate("plant_vitoria")]
        for planta in plantas:
            radio = ctk.CTkRadioButton(grupo_planta, text=planta, variable=self.planta_selecionada, value=planta, font=FONTE_PADRAO, command=self.verificar_selecao)
            radio.pack(anchor="w", pady=2, padx=10)
            self.plant_radiobuttons.append(radio)

        # Interval Selection Group Frame
        grupo_intervalo = ctk.CTkFrame(painel_direito, corner_radius=6)
        grupo_intervalo.pack(pady=10, padx=10, fill="x")

        # Label for interval group
        self.rotulo_intervalo = ctk.CTkLabel(grupo_intervalo, text=self.localizer.translate("record_interval"), font=FONTE_SUBTITULO)
        self.rotulo_intervalo.pack(pady=(0, 5), padx=10, anchor="w")

        # Interval Radio Buttons
        self.interval_radiobuttons = []
        intervalos = [
            (self.localizer.translate("30_seconds"), "30 segundos"),
            (self.localizer.translate("1_minute"), "1 minuto")
        ]
        for texto, valor in intervalos:
            radio = ctk.CTkRadioButton(grupo_intervalo, text=texto, variable=self.intervalo_selecionado, value=valor, font=FONTE_PADRAO, command=self.verificar_selecao)
            radio.pack(anchor="w", pady=2, padx=10)
            self.interval_radiobuttons.append(radio)

        # Analysis Duration Selection Group Frame
        grupo_duracao = ctk.CTkFrame(painel_direito, corner_radius=6)
        grupo_duracao.pack(pady=10, padx=10, fill="x")

        # Label for duration group
        self.rotulo_duracao = ctk.CTkLabel(grupo_duracao, text=self.localizer.translate("analysis_duration"), font=FONTE_SUBTITULO)
        self.rotulo_duracao.pack(pady=(0, 5), padx=10, anchor="w")

        # Duration Radio Buttons
        self.duration_radiobuttons = []
        duracoes = [
            (self.localizer.translate("30_seconds"), "30 segundos"),
            (self.localizer.translate("2_minutes"), "2 minutos"),
            (self.localizer.translate("3_minutes"), "3 minutos"),
            (self.localizer.translate("5_minutes"), "5 minutos"),
            (self.localizer.translate("10_minutes"), "10 minutos"),
            (self.localizer.translate("20_minutes"), "20 minutos"),
            (self.localizer.translate("30_minutes"), "30 minutos"),
            (self.localizer.translate("35_minutes"), "35 minutos")
        ]
        for texto, valor in duracoes:
            radio = ctk.CTkRadioButton(grupo_duracao, text=texto, variable=self.analise_duracao_selected, value=valor, font=FONTE_PADRAO, command=self.verificar_selecao)
            radio.pack(anchor="w", pady=2, padx=10)
            self.duration_radiobuttons.append(radio)

        # Sample ID Entry Fields
    
        campos_id_frame = ctk.CTkFrame(painel_direito, corner_radius=6)
        campos_id_frame.pack(pady=20, padx=10, fill="x")

        self.rotulo_id = ctk.CTkLabel(campos_id_frame, text=self.localizer.translate("sample_ids"), font=FONTE_SUBTITULO)
        self.rotulo_id.pack(pady=(0, 5), padx=10, anchor="w")

        self.sample_id_labels = {}
        for termopar in ["T1", "T2", "T3", "T4"]:
            frame = ctk.CTkFrame(campos_id_frame, corner_radius=6)
            frame.pack(pady=2, padx=10, fill="x")
            
            label = ctk.CTkLabel(frame, text=f"{self.localizer.translate('id')} {termopar}:", font=FONTE_PADRAO)
            label.pack(side="left", padx=5, pady=5)
            self.sample_id_labels[termopar] = label

            entry = ctk.CTkEntry(frame, textvariable=self.codigos_amostras_vars[termopar], font=FONTE_PADRAO)
            entry.pack(side="left", padx=5, pady=5, fill="x", expand=True)
            entry.configure(state="disabled")  # Initially disabled
            self.entry_widgets[termopar] = entry  # Store entry widget
    
    def create_main_frame(self):
        """
        Creates the main frame divided into visualization area.
        """
        main_frame = ctk.CTkFrame(self, corner_radius=6)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Visualization Area
        area_visualizacao = ctk.CTkFrame(main_frame, corner_radius=6)
        area_visualizacao.pack(fill="both", expand=True)

        # Use grid layout to control the placement of panels
        area_visualizacao.grid_rowconfigure(0, weight=1)
        area_visualizacao.grid_columnconfigure(0, weight=1)
        area_visualizacao.grid_columnconfigure(1, weight=1)

        # Frames for each visualization panel
        self.painel_grafico = ctk.CTkFrame(area_visualizacao, corner_radius=6)
        self.painel_tabela = ctk.CTkFrame(area_visualizacao, corner_radius=6)
        
        self.painel_grafico.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.painel_tabela.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        # Initialize Panels
        self.create_painel_grafico()
        self.create_painel_tabela()

    def create_painel_grafico(self):
        """
        Creates the graph panel with real-time temperatures.
        Includes a Matplotlib graph and display of each thermocouple's temperatures.
        """
        # Time vs Temperature Graph
        self.fig, self.ax = plt.subplots(figsize=(6, 4))
        self.ax.set_title("Curva de Reatividade da Amostra", fontsize=16)
        self.ax.set_xlabel("Tempo (s)", fontsize=12)
        self.ax.set_ylabel("Temperatura (°C)", fontsize=12)
        self.ax.set_xlim(0, 300)  # Dynamically adjust
        self.ax.set_ylim(20, 100)  # Adjust as needed

        # Define colors for the graph lines
        cores = ["#e6194b", "#3cb44b", "#ffe119", "#4363d8"]  # Example colors

        self.lines = {}
        for i, termopar in enumerate(["T1", "T2", "T3", "T4"]):
            line, = self.ax.plot([], [], label=f"{termopar} (°C)", color=cores[i % len(cores)], linewidth=2)
            self.lines[termopar] = line
        self.ax.legend(fontsize=10)

        # Graph Canvas
        self.canvas_grafico = FigureCanvasTkAgg(self.fig, master=self.painel_grafico)
        self.canvas_grafico.draw()
        self.canvas_grafico.get_tk_widget().pack(pady=10, padx=10, fill="both", expand=True)

        # Space to show real-time temperatures (Changed to Grid)
        self.temp_labels = {}
        temp_frame = ctk.CTkFrame(self.painel_grafico)
        temp_frame.pack(pady=5, padx=5, anchor="center", fill="x")

        # Arrange labels in grid (2 columns)
        for index, termopar in enumerate(["T1", "T2", "T3", "T4"]):
            row = index // 2
            col = index % 2
            label = ctk.CTkLabel(temp_frame, text=f"{termopar}: 0.00°C", font=("Roboto", 16, "bold"))
            label.grid(row=row, column=col, padx=20, pady=10, sticky="nsew")
            self.temp_labels[termopar] = label

        # Configure column expansion so labels occupy available space
        temp_frame.columnconfigure(0, weight=1)
        temp_frame.columnconfigure(1, weight=1)

    def create_painel_tabela(self):
        """
        Creates the table panel that will be filled as the analysis progresses.
        Uses a ttk.Treeview for data display.
        """
        # Data Table using ttk.Treeview
        self.table = ttk.Treeview(self.painel_tabela, columns=("Tempo", "T1", "T2", "T3", "T4"), show='headings')
        self.table.heading("Tempo", text="Tempo")
        self.table.heading("T1", text="T1 (°C)")
        self.table.heading("T2", text="T2 (°C)")
        self.table.heading("T3", text="T3 (°C)")
        self.table.heading("T4", text="T4 (°C)")

        # Set column widths
        self.table.column("Tempo", width=80, anchor="center")
        self.table.column("T1", width=80, anchor="center")
        self.table.column("T2", width=80, anchor="center")
        self.table.column("T3", width=80, anchor="center")
        self.table.column("T4", width=80, anchor="center")

        # Add a scrollbar
        scrollbar = ttk.Scrollbar(self.painel_tabela, orient="vertical", command=self.table.yview)
        self.table.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.table.pack(pady=10, padx=10, fill="both", expand=True)

    def create_bottom_frame(self):
        """
        Creates the bottom area with progress bar, status, and button to interrupt analysis.
        Initially, the progress bar and button are hidden.
        """
        self.bottom_frame = ctk.CTkFrame(self, corner_radius=6)
        self.bottom_frame.pack(side="bottom", fill="x", padx=10, pady=10)

        # Progress Bar (Initially Hidden)
        self.progress_bar = ctk.CTkProgressBar(self.bottom_frame, width=400, height=20)
        self.progress_bar.set(0)
        # Do not pack here

        # Status Label
        self.status_label = ctk.CTkLabel(self.bottom_frame, text=f"{self.localizer.translate('status')}: {self.localizer.translate('waiting')}", font=FONTE_PADRAO)
        self.status_label.pack(side="left", padx=20)

        # Button to Interrupt Analysis (Initially Hidden)
        self.btn_interromper = ctk.CTkButton(
            self.bottom_frame,
            text=self.localizer.translate("interrupt_analysis"),
            command=self.interromper_analise,
            fg_color="#FF4D4D",
            hover_color="#FF6666",
            font=FONTE_PADRAO
        )
        # Do not pack here

    def verificar_selecao(self, *args):
        """
        Checks if all RadioButton options are selected to enable the start button.
        """
        pass  # Start button is always enabled

    def verificar_termopares(self):
        """
        Checks active thermocouples and enables the corresponding text boxes.
        """
        if self.simulacao_ativa:
            # Simulation: randomly select active thermocouples
            self.termopares_ativos = [tp for tp in ["T1", "T2", "T3", "T4"] if random.choice([True, False])]
            logger.info(f"Simulação: Termopares ativos detectados: {self.termopares_ativos}")
        else:
            # Real mode: identify active thermocouples via serial port
            manipulador_serial = ManipuladorPortaSerial("COM12")  # Replace "COM12" with the correct port
            if not manipulador_serial.abrir():
                messagebox.showerror(self.localizer.translate("error"), self.localizer.translate("cannot_open_serial_port"))
                return

            canal_para_termopar = {"41": "T1", "42": "T2", "43": "T3", "44": "T4"}
            self.termopares_ativos = identificar_termopares_ativos(manipulador_serial, canal_para_termopar)
            manipulador_serial.fechar()

            if not self.termopares_ativos:
                messagebox.showwarning(self.localizer.translate("attention"), self.localizer.translate("no_active_thermocouples_detected"))
                self.termopares_ativos = []
            else:
                logger.info(f"Termopares ativos detectados: {self.termopares_ativos}")

        # Enable or disable text boxes based on active thermocouples
        for termopar, var in self.codigos_amostras_vars.items():
            entry_widget = self.entry_widgets.get(termopar)
            if termopar in self.termopares_ativos:
                # Enable entry
                if entry_widget:
                    entry_widget.configure(state="normal")
            else:
                # Disable and clear entry
                if entry_widget:
                    entry_widget.configure(state="disabled")
                    var.set("")  # Clear the field

    def iniciar_analise(self):
        """
        Starts the analysis after checking selections and confirms with the user.
        """
        if not self.analise_em_andamento:
            # Check if all IDs for active thermocouples are filled
            faltando_ids = [tp for tp in self.termopares_ativos if not self.codigos_amostras_vars[tp].get().strip()]
            if faltando_ids:
                messagebox.showwarning(self.localizer.translate("incomplete_fields"), f"{self.localizer.translate('please_fill_ids')} {', '.join(faltando_ids)}.")
                return

            # Confirm with the user before starting
            nomes_amostras = ", ".join([f"ID {tp}: {var.get().strip()}" for tp, var in self.codigos_amostras_vars.items() if tp in self.termopares_ativos])
            msg = (f"{self.localizer.translate('about_to_start_analysis')}:\n{nomes_amostras}.\n\n"
                   f"{self.localizer.translate('click_ok_to_start')}\n"
                   f"{self.localizer.translate('click_cancel_to_abort')}.")
            resposta = messagebox.askokcancel(self.localizer.translate("confirm_start"), msg)
            if resposta:
                logger.info("Usuário confirmou o início da medição.")
                self.analise_em_andamento = True
                self.btn_iniciar.configure(state="disabled")
                self.btn_exportar.configure(state="disabled")
                self.status_label.configure(text=f"{self.localizer.translate('status')}: {self.localizer.translate('analyzing')}...")

                # Reset graphs and tables before starting
                self.resetar_grafico()
                self.resetar_tabela()

                # Display the progress bar and interrupt button
                self.progress_bar.pack(side="left", padx=20, pady=10, fill="x", expand=True)
                self.btn_interromper.pack(side="right", padx=20)

                self.iniciar_medicao()
            else:
                logger.info("Usuário cancelou o início da medição.")

    def iniciar_medicao(self):
        """
        Starts the measurement in a separate thread.
        """
        # Reset previous data
        self.dados = []
        # Reset graph and table before starting
        self.resetar_grafico()
        self.resetar_tabela()

        # Collect sample IDs
        codigos_amostras = {tp: var.get().strip() for tp, var in self.codigos_amostras_vars.items() if tp in self.termopares_ativos}

        # Initialize TemperatureLogger or SimulatedTemperatureLogger
        if self.simulacao_ativa:
            self.temp_logger_simulado = SimulatedTemperatureLogger(self.termopares_ativos)
            self.temp_logger_simulado.start()
            temperature_logger_instance = self.temp_logger_simulado
        else:
            # Use the real TemperatureLogger
            manipulador_serial = ManipuladorPortaSerial("COM12")  # Replace "COM12" with the correct port
            if not manipulador_serial.abrir():
                messagebox.showerror(self.localizer.translate("error"), self.localizer.translate("cannot_open_serial_port"))
                self.analise_em_andamento = False
                self.btn_iniciar.configure(state="normal")
                self.btn_exportar.configure(state="disabled")
                self.status_label.configure(text=f"{self.localizer.translate('status')}: {self.localizer.translate('waiting')}")
                return
            canal_para_termopar = {"41": "T1", "42": "T2", "43": "T3", "44": "T4"}
            temperature_logger_instance = TemperatureLogger(manipulador_serial, self.termopares_ativos, canal_para_termopar)
            temperature_logger_instance.start()

        # Define measurement duration based on selection
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
            # If no match, set default values
            tempo_total = 30

        self.tempo_total = tempo_total  # Store total time for progress bar use

        # Get the recording interval
        intervalo = self.intervalo_selecionado.get()
        if intervalo == "30 segundos":
            intervalo_segundos = 30
        elif intervalo == "1 minuto":
            intervalo_segundos = 60
        else:
            intervalo_segundos = 30  # Default value

        # Calculate all recording times
        log_times = []
        current_time = 0
        while current_time <= tempo_total:
            log_times.append(current_time)
            current_time += intervalo_segundos

        # Ensure tempo_total is included
        if log_times[-1] != tempo_total:
            log_times.append(tempo_total)

        # Sort and remove duplicates
        tempos_registro = sorted(list(set(log_times)))

        # Set the mode name
        nome_modo = "Simulação" if self.simulacao_ativa else "Real"

        # Instantiate Medicao with all required arguments
        self.medicao = Medicao(
            temperature_logger_instance,
            codigos_amostras,
            tempos_registro,
            tempo_total,
            nome_modo
        )

        # Start the measurement thread
        self.thread_medicao = threading.Thread(target=self.run_medicao, args=(tempo_total, tempos_registro))
        self.thread_medicao.daemon = True  # Allows the thread to be terminated with the application
        self.thread_medicao.start()

    def run_medicao(self, tempo_total, log_times):
        try:
            start_time = time.time()
            last_update_time = -1
            while True:
                if not self.analise_em_andamento:
                    break

                elapsed_time = int(time.time() - start_time)

                # Update temperatures every second
                if elapsed_time != last_update_time:
                    last_update_time = elapsed_time
                    temperaturas = self.medicao.obter_temperaturas()
                    # Schedule GUI update
                    self.after(0, self.atualizar_gui, elapsed_time, temperaturas)
                    # Append data for plotting
                    self.dados.append((elapsed_time, temperaturas.copy()))

                # Record temperatures at specified intervals
                if elapsed_time in log_times:
                    # Log the temperature
                    temperaturas = self.medicao.obter_temperaturas()
                    logger.info(f"Temperature record at {formatar_tempo(elapsed_time)}: {temperaturas}")

                if elapsed_time >= tempo_total:
                    break

                time.sleep(0.1)

            # Finish the measurement
            self.delta_t = self.calcular_delta_t()
            self.after(0, self.atualizar_delta_t_sidebar, self.delta_t)
            self.after(0, self.status_label.configure, {'text': f"{self.localizer.translate('status')}: {self.localizer.translate('completed')}"})
            self.after(0, lambda: messagebox.showinfo(self.localizer.translate("measurement"), self.localizer.translate("measurement_completed_successfully")))
            logger.info("Measurement completed successfully.")
        except Exception as e:
            logger.error(f"Error during measurement: {e}")
            self.after(0, lambda: messagebox.showerror(self.localizer.translate("error"), f"{self.localizer.translate('error_during_measurement')}\n{e}"))
            self.after(0, self.status_label.configure, {'text': f"{self.localizer.translate('status')}: {self.localizer.translate('measurement_error')}"})
        finally:
            self.analise_em_andamento = False
            # Stop the TemperatureLogger
            if self.simulacao_ativa and hasattr(self, 'temp_logger_simulado'):
                self.temp_logger_simulado.parar()
            elif hasattr(self.medicao, 'temp_logger'):
                self.medicao.temp_logger.parar()
            # Hide the progress bar and interrupt button
            self.after(0, self.progress_bar.pack_forget)
            self.after(0, self.btn_interromper.pack_forget)
            # Enable the export button
            self.after(0, self.btn_exportar.configure, {'state': "normal"})

    def atualizar_delta_t_sidebar(self, delta_t):
        """
        Updates the Delta T label in the left sidebar.
        """
        self.delta_label_sidebar.configure(text=f"{self.localizer.translate('delta_t')}: {delta_t:.2f} °C")

    def calcular_delta_t(self):
        """
        Calculates Delta T based on collected data.
        """
        if self.dados:
            temperaturas_iniciais = []
            temperaturas_finais = []
            for temp in self.dados[0][1].values():
                try:
                    temperaturas_iniciais.append(float(temp))
                except ValueError:
                    temperaturas_iniciais.append(0.0)
            for temp in self.dados[-1][1].values():
                try:
                    temperaturas_finais.append(float(temp))
                except ValueError:
                    temperaturas_finais.append(0.0)
            delta = sum([f - i for i, f in zip(temperaturas_iniciais, temperaturas_finais)]) / len(temperaturas_iniciais)
            return delta
        return 0.0

    def exportar_dados(self):
        """
        Exports the measurement data to an Excel file.
        """
        try:
            if not self.dados:
                messagebox.showwarning(self.localizer.translate("export"), self.localizer.translate("no_data_to_export"))
                return

            # Initialize ExportadorDados
            exportador = ExportadorDados(self.dados, self.codigos_amostras_vars, tipo_analise="comum")
            caminho_exportacao = exportador.exportar_para_excel()

            messagebox.showinfo(self.localizer.translate("export"), f"{self.localizer.translate('data_exported_successfully')} {caminho_exportacao}")
            logger.info(f"Dados exportados para {caminho_exportacao}")
        except Exception as e:
            logger.error(f"Erro ao exportar dados: {e}", exc_info=True)
            messagebox.showerror(self.localizer.translate("error"), f"{self.localizer.translate('error_exporting_data')}\n{e}")

    def resetar_campos(self):
        """
        Resets all selections, entries, and interface states.
        """
        self.planta_selecionada.set("São José da Lapa")
        self.intervalo_selecionado.set("30 segundos")
        self.analise_duracao_selected.set("30 segundos")
        self.btn_iniciar.configure(state="normal")
        self.btn_exportar.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text=f"{self.localizer.translate('status')}: {self.localizer.translate('waiting')}")
        self.delta_label_sidebar.configure(text=f"{self.localizer.translate('delta_t')}: N/A")

        # Hide the progress bar and interrupt button if visible
        self.progress_bar.pack_forget()
        self.btn_interromper.pack_forget()

        # Reset sample ID entries
        for termopar, var in self.codigos_amostras_vars.items():
            entry_widget = self.entry_widgets.get(termopar)
            if entry_widget:
                entry_widget.configure(state="disabled")
                var.set("")

        # Reset data
        self.dados = []

        # Reset graphs and table
        self.resetar_grafico()
        self.resetar_tabela()

        # Reset temperatures
        for termopar, label in self.temp_labels.items():
            label.configure(text=f"{termopar}: 0.00°C")

        logger.info("Campos resetados pelo usuário.")

    def interromper_analise(self):
        """
        Interrupts the ongoing analysis.
        """
        if self.analise_em_andamento:
            resposta = messagebox.askyesno(self.localizer.translate("interrupt_analysis"), self.localizer.translate("confirm_interrupt_analysis"))
            if resposta:
                self.analise_em_andamento = False
                if self.simulacao_ativa and hasattr(self, 'temp_logger_simulado'):
                    self.temp_logger_simulado.parar()
                elif hasattr(self.medicao, 'temp_logger'):
                    self.medicao.temp_logger.parar()
                logger.info("Análise interrompida pelo usuário.")
                messagebox.showinfo(self.localizer.translate("interrupt_analysis"), self.localizer.translate("analysis_interrupted_by_user"))
                # Reset progress bar and status
                self.progress_bar.set(0)
                self.status_label.configure(text=f"{self.localizer.translate('status')}: {self.localizer.translate('analysis_interrupted')}")
                # Hide the progress bar and interrupt button
                self.progress_bar.pack_forget()
                self.btn_interromper.pack_forget()
                # Enable the export button
                self.btn_exportar.configure(state="normal")

    def atualizar_gui(self, tempo, temperaturas):
        """
        Updates the graph, table, progress bar, and real-time temperatures.
        """
        # Update Real-Time Temperatures
        for termopar in self.termopares_ativos:
            temp = temperaturas.get(termopar, 0.0)
            try:
                temp_float = float(temp)
            except (ValueError, TypeError):
                temp_float = 0.0
            self.temp_labels[termopar].configure(text=f"{termopar}: {temp_float:.2f}°C")

        # Update Graph
        tempos = [d[0] for d in self.dados]
        for termopar in self.termopares_ativos:
            temps = []
            for d in self.dados:
                temp = d[1].get(termopar, 0)
                try:
                    temp_float = float(temp)
                except (ValueError, TypeError):
                    temp_float = 0.0
                temps.append(temp_float)
            self.lines[termopar].set_data(tempos, temps)
        self.ax.set_xlim(0, max(self.tempo_total, max(tempos) + 10))
        self.ax.relim()
        self.ax.autoscale_view()
        self.canvas_grafico.draw()

        # Update Progress Bar
        percent = (tempo / self.tempo_total) * 100 if self.tempo_total > 0 else 0
        self.progress_bar.set(percent / 100)

        # Update Table
        tempo_str = formatar_tempo(tempo)
        temp_values = []
        for tp in ["T1", "T2", "T3", "T4"]:
            temp = temperaturas.get(tp, 0)
            try:
                temp_float = float(temp)
                temp_str = f"{temp_float:.2f}"
            except (ValueError, TypeError):
                temp_str = "N/A"
            temp_values.append(temp_str)
        self.table.insert("", "end", values=(tempo_str, *temp_values))

    def resetar_grafico(self):
        """
        Resets the graph to an empty state.
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
        Resets the table to an empty state.
        """
        for item in self.table.get_children():
            self.table.delete(item)

    def on_closing(self):
        """
        Function called when the window is being closed.
        Ensures that all threads are properly terminated.
        """
        if self.analise_em_andamento:
            resposta = messagebox.askyesno(self.localizer.translate("exit"), self.localizer.translate("exit_confirmation_analysis"))
            if resposta:
                self.analise_em_andamento = False
                if self.simulacao_ativa and hasattr(self, 'temp_logger_simulado'):
                    self.temp_logger_simulado.parar()
                elif hasattr(self.medicao, 'temp_logger'):
                    self.medicao.temp_logger.parar()
                self.destroy()
            else:
                return
        else:
            if messagebox.askyesno(self.localizer.translate("exit"), self.localizer.translate("exit_confirmation")):
                self.destroy()
                sys.exit()

if __name__ == "__main__":
    # Set the appearance mode to 'light' before starting the application
    ctk.set_appearance_mode("light")

    # Check if the custom theme exists
    tema_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "themes", "blue-light.json")
    if not os.path.exists(tema_path):
        messagebox.showerror("Erro", "Arquivo de tema 'blue-light.json' não encontrado na pasta 'themes'.")
    else:
        app = ReactLabApp()
        app.mainloop()
