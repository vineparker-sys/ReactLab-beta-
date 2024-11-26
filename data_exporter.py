import os
import pandas as pd
import logging
import xlsxwriter
import datetime
import ctypes
from ctypes import wintypes
import string

class ExportadorDados:
    """Exporta os dados coletados para um arquivo Excel, incluindo formatações e cálculos, e protege o arquivo com senha."""

    def __init__(self, dados, codigos_amostras, tipo_analise):
        """Inicializa o exportador com os dados, códigos das amostras e tipo de análise."""
        self.dados = dados
        self.codigos_amostras = codigos_amostras
        self.tipo_analise = tipo_analise

    def formatar_tempo_excel(self, segundos):
        """Formata o tempo em mm:ss para exibição no Excel."""
        minutos = segundos // 60
        segundos_restantes = segundos % 60
        return f"{int(minutos):02d}:{int(segundos_restantes):02d}"

    def exportar_para_excel(self):
        """Exporta os dados para um arquivo Excel, aplicando formatações, calculando o delta T e protegendo a planilha com senha."""
        # Determinar termopares ativos
        termopares_ativos = set()
        for _, temperaturas in self.dados:
            for termopar, temp in temperaturas.items():
                termopares_ativos.add(termopar)

        if not termopares_ativos:
            print("Nenhum termopar ativo foi detectado para exportação.")
            return

        # Criar DataFrame com o tempo formatado
        df = pd.DataFrame({"Tempo": [self.formatar_tempo_excel(t[0]) for t in self.dados]})

        # Adicionar as temperaturas ao DataFrame
        for termopar in sorted(termopares_ativos):
            df[termopar] = [
                float(dado[1][termopar].split()[0]) if dado[1][termopar] != 'OFF' else None
                for dado in self.dados
            ]

        # Nome do arquivo com o código da amostra e timestamp
        # Se houver apenas um código de amostra, usamos ele; caso contrário, indicamos 'Multiplas_Amostras'
        codigos = set(self.codigos_amostras.values())
        if len(codigos) == 1:
            codigo_amostra = codigos.pop()
        else:
            codigo_amostra = "Multiplas_Amostras"

        # Sanitizar o código da amostra para remover caracteres inválidos
        valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
        codigo_amostra_sanitizado = ''.join(c for c in codigo_amostra if c in valid_chars)

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        # Determinar o caminho para a área de trabalho do usuário de forma confiável
        CSIDL_DESKTOP = 0
        SHGFP_TYPE_CURRENT = 0

        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOP, None, SHGFP_TYPE_CURRENT, buf)
        desktop_path = buf.value

        # Nome completo do arquivo a ser salvo na área de trabalho
        file_path = os.path.join(desktop_path, f"Amostra_{codigo_amostra_sanitizado}_{timestamp}.xlsx")

        try:
            # Exportar para Excel com XlsxWriter usando o contexto gerenciado
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Dados', index=False, startrow=1, header=False)

                workbook = writer.book
                worksheet = writer.sheets['Dados']

                # Proteger a planilha com senha (somente para edição)
                worksheet.protect(password='Lhoist@2024!')

                # Formatação das células
                header_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})
                cell_format = workbook.add_format({'border': 1})
                blue_cell_format = workbook.add_format({'border': 1, 'bg_color': '#BDD7EE'})
                time_format = workbook.add_format({'border': 1})

                # Ajustar largura das colunas
                worksheet.set_column('A:A', 10)  # Tempo
                worksheet.set_column('B:Z', 15)  # Termopares

                # Escrever cabeçalhos com formatação
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(1, col_num, value, header_format)

                # Escrever os dados com formatação
                for row_num in range(len(df)):
                    for col_num in range(len(df.columns)):
                        # Identificar T0 e Tf para aplicar o fundo azul
                        if row_num == 0 or row_num == len(df) - 1:
                            cell_fmt = blue_cell_format
                        else:
                            cell_fmt = cell_format

                        cell_value = df.iloc[row_num, col_num]
                        worksheet.write(row_num + 2, col_num, cell_value, cell_fmt)

                # Calcular o delta T e inserir abaixo da tabela
                delta_row = len(df) + 3  # Linha onde o delta T será inserido
                worksheet.write(delta_row, 0, "Delta T (°C)", header_format)

                for termopar in sorted(termopares_ativos):
                    col_num = df.columns.get_loc(termopar)
                    initial_temp = df[termopar].iloc[0]
                    final_temp = df[termopar].iloc[-1]
                    if initial_temp is not None and final_temp is not None:
                        delta_t = final_temp - initial_temp
                        worksheet.write(delta_row, col_num, delta_t, cell_format)
                    else:
                        worksheet.write(delta_row, col_num, "N/A", cell_format)

                # Inserir gráficos para cada termopar à direita da tabela
                chart_col_position = len(df.columns) + 1  # Coluna após os dados
                chart_row_position = 1  # Linha inicial para o gráfico

                for termopar in sorted(termopares_ativos):
                    chart = workbook.add_chart({'type': 'line'})
                    chart.add_series({
                        'name': termopar,
                        'categories': ['Dados', 2, 0, len(df) + 1, 0],  # Tempo
                        'values': ['Dados', 2, df.columns.get_loc(termopar), len(df) + 1, df.columns.get_loc(termopar)],
                    })
                    chart.set_title({'name': f'Curva de Temperatura - {termopar}'})
                    chart.set_x_axis({'name': 'Tempo'})
                    chart.set_y_axis({'name': 'Temperatura (°C)'})
                    chart.set_size({'width': 600, 'height': 400})

                    # Inserir o gráfico à direita da tabela, com espaço entre gráficos
                    worksheet.insert_chart(chart_row_position, chart_col_position, chart)

                    # Ajustar a posição da linha para o próximo gráfico (espaço de 20 linhas entre gráficos)
                    chart_row_position += 20  # Ajuste conforme necessário

            print(f"Dados exportados com sucesso para {file_path}!")
            logging.info(f"Dados exportados para {file_path}")
            return file_path
        except Exception as e:
            print(f"Erro ao exportar para Excel: {e}")
            logging.error(f"Erro ao exportar para Excel: {e}")
