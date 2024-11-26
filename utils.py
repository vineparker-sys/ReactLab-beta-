# utils.py

import re
import logging
from rich.console import Console
from rich.table import Table

console = Console()

def limpar_tela():
    """Limpa o console."""
    console.clear()

def formatar_tempo(segundos):
    """Formata o tempo em minutos e segundos."""
    minutos = int(segundos // 60)
    segundos_restantes = int(segundos % 60)
    return f"{minutos}:{segundos_restantes:02d}"

def exibir_cabecalho():
    """Retorna o cabeçalho estilizado para exibição no console."""
    header = """
[bold blue]
██╗      ██╗  ██╗ ██████╗ ██╗███████╗████████╗
██║      ██║  ██║██╔═══██╗██║██╔════╝╚══██╔══╝
██║      ███████║██║   ██║██║███████╗   ██║   
██║      ██╔══██║██║   ██║██║╚════██║   ██║   
███████╗ ██║  ██║╚██████╔╝██║███████║   ██║   
╚══════╝ ╚═╝  ╚═╝ ╚═════╝ ╚═╝╚══════╝   ╚═╝  
[/bold blue]
"""
    return header

def criar_tabela_temperaturas(temperaturas, codigos_amostras):
    """Cria uma tabela para exibição das temperaturas dos termopares."""
    table = Table(show_header=True, header_style="bold blue")
    table.add_column("Termopar", justify="center", style="white")
    table.add_column("Temperatura", justify="center", style="white")
    table.add_column("Código Amostra", justify="center", style="white")

    for termopar, temp in temperaturas.items():
        codigo_amostra = codigos_amostras.get(termopar, "N/A")
        if temp != 'OFF':
            table.add_row(f"[green]{termopar}[/green]", f"[green]{temp}[/green]", f"[green]{codigo_amostra}[/green]")
        else:
            table.add_row(f"[red]{termopar}[/red]", f"[red]{temp}[/red]", f"[red]{codigo_amostra}[/red]")

    return table

def extrair_temperaturas(dados, canal_para_termopar):
    """Extrai as temperaturas dos dados brutos lidos."""
    termopares_novos = {}
    matches = re.findall(r'(\d{10,14})', dados)
    logger = logging.getLogger("utils")
    for match in matches:
        if len(match) == 14:
            canal = match[:2]
            temperatura_bruta = match[-3:]
            try:
                temperatura = int(temperatura_bruta[:-1]) + (int(temperatura_bruta[-1]) / 10.0)
                temperatura_formatada = f"{temperatura:.1f} °C"
            except ValueError:
                logger.warning(f"Erro ao converter temperatura bruta '{temperatura_bruta}' para canal '{canal}'")
                continue

            if canal in canal_para_termopar:
                termopar = canal_para_termopar[canal]
                termopares_novos[termopar] = temperatura_formatada

    if not termopares_novos:
        logger.debug("Nenhum valor de temperatura extraído dos dados.")
    return termopares_novos if termopares_novos else None
