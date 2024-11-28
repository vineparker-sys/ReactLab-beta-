# data_exporter.py

import os
import pandas as pd
import xlsxwriter
from utils import formatar_tempo
import getpass  # Para obter o nome do usuário
import numpy as np

class ExportadorDados:
    def __init__(self, dados, codigos_amostras_vars, intervalo_registro, planta_selecionada, tipo_analise="comum"):
        self.dados = dados  # Lista de tuplas (elapsed_time, temperatures dict)
        self.codigos_amostras_vars = codigos_amostras_vars  # Dicionário de IDs das amostras
        self.intervalo_registro = intervalo_registro  # Intervalo de registro em segundos
        self.planta_selecionada = planta_selecionada  # Nome da planta selecionada
        self.tipo_analise = tipo_analise

    def exportar_para_excel(self, save_path=None, filename=None):
        """
        Exporta os dados para um arquivo Excel.
        """
        # Verificar termopares ativos com IDs de amostras
        active_termopares = [tp for tp, id_amostra in self.codigos_amostras_vars.items() if id_amostra.strip()]
        if not active_termopares:
            raise ValueError("Nenhum termopar ativo com IDs de amostras.")

        # Construir um DataFrame com Tempo e Temperaturas
        data = {'Tempo': []}
        for tp in active_termopares:
            data[tp] = []

        # Filtrar os dados para incluir apenas os tempos no intervalo de registro
        tempos_registro = set()
        current_time = 0
        tempo_total = self.dados[-1][0]  # Último tempo registrado
        while current_time <= tempo_total:
            tempos_registro.add(current_time)
            current_time += self.intervalo_registro

        # Garantir que o tempo 0 esteja incluído
        tempos_registro.add(0)

        for elapsed_time, temps_dict in self.dados:
            if elapsed_time in tempos_registro:
                data['Tempo'].append(formatar_tempo(elapsed_time))
                for tp in active_termopares:
                    temp = temps_dict.get(tp, None)
                    if temp is not None:
                        try:
                            temp_float = float(temp)
                            data[tp].append(temp_float)
                        except ValueError:
                            data[tp].append(np.nan)
                    else:
                        data[tp].append(np.nan)

        df = pd.DataFrame(data)

        # Calcular Delta T para cada termopar
        delta_ts = {}
        for tp in active_termopares:
            delta_ts[tp] = self.calcular_delta_t(df[tp].values)

        # Obter o nome do usuário logado
        nome_usuario = getpass.getuser()

        # Definir o caminho e o nome do arquivo
        if save_path is None:
            save_path = os.getcwd()
        if filename is None:
            filename = 'Resultados_Medicao.xlsx'
        filepath = os.path.join(save_path, filename)

        # Criar um escritor do Excel com opção para tratar NaN/Inf
        with pd.ExcelWriter(filepath, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
            # Escrever DataFrame no Excel
            df.to_excel(writer, index=False, sheet_name='Dados', startrow=1)

            # Obter o workbook e worksheet
            workbook = writer.book
            worksheet = writer.sheets['Dados']

            # Formatações
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
            cell_format = workbook.add_format({'align': 'center', 'border': 1})
            delta_format = workbook.add_format({'align': 'left', 'border': 1, 'bold': True})
            info_format = workbook.add_format({'align': 'left', 'border': 1})

            # Aplicar formatação nos cabeçalhos
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(1, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 15)

            # Aplicar formatação nas células de dados
            for row_num in range(2, len(df) + 2):
                for col_num in range(len(df.columns)):
                    val = df.iloc[row_num - 2, col_num]
                    if pd.isna(val):
                        val = ''
                    worksheet.write(row_num, col_num, val, cell_format)

            # Escrever Delta T abaixo da tabela
            delta_t_row = len(df) + 3  # Ajuste para posicionar corretamente
            worksheet.write(delta_t_row, 0, 'Delta T:', delta_format)
            col = 1
            for tp in active_termopares:
                worksheet.write(delta_t_row, col, f'{tp}: {delta_ts[tp]} °C', delta_format)
                col += 1

            # Escrever informações do usuário e planta
            info_row = delta_t_row + 2
            info_text = f"Análise realizada pelo usuário: '{nome_usuario}' - planta '{self.planta_selecionada}'"
            worksheet.write(info_row, 0, info_text, info_format)

            # Ajustar largura das colunas para acomodar o texto
            worksheet.set_column(0, 0, 50)

            # Proteger a planilha com senha
            worksheet.protect('LH015Tl4b!')

            # Criar um gráfico
            chart = workbook.add_chart({'type': 'line'})
            for tp in active_termopares:
                col_idx = df.columns.get_loc(tp)
                chart.add_series({
                    'categories': ['Dados', 2, 0, len(df) + 1, 0],  # Coluna Tempo
                    'values':     ['Dados', 2, col_idx, len(df) + 1, col_idx],  # Coluna do termopar
                    'name':       f'{tp} - {self.codigos_amostras_vars[tp]}'
                })
            chart.set_title({'name': 'Curva de Reatividade'})
            chart.set_x_axis({'name': 'Tempo'})
            chart.set_y_axis({'name': 'Temperatura (°C)'})

            # Inserir gráfico na worksheet, ao lado direito da tabela
            chart_start_col = len(df.columns) + 2  # Coluna após os dados + espaçamento
            worksheet.insert_chart(1, chart_start_col, chart)

            # Ajustar o tamanho do gráfico para não sobrepor os dados
            chart.set_size({'width': 600, 'height': 400})

            # Ajustar zoom da planilha
            worksheet.set_zoom(90)

        return os.path.abspath(filepath)

    def calcular_delta_t(self, temperatures):
        """
        Calcula o Delta T com base na lista de temperaturas.
        """
        temperaturas_validas = [temp for temp in temperatures if not np.isnan(temp)]
        if temperaturas_validas:
            initial_temp = temperaturas_validas[0]
            final_temp = temperaturas_validas[-1]
            delta_t = final_temp - initial_temp
            return round(delta_t, 2)
        else:
            return 0.0
