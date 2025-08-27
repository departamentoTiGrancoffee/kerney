import pandas as pd
import xlwings as xw
import os
import time
import win32com.client

def close_excel_file_if_open(filename):
    """Check if an Excel file is open and close it using pywin32."""
    filename = os.path.basename(filename)  # Extract just the filename
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")  # Get running Excel
        for wb in excel.Workbooks:
            if wb.Name == filename:
                print(f"Closing {filename}...")
                wb.Close(SaveChanges=False)  # Close without saving
                time.sleep(1)  # Give Excel a moment
                return True
    except Exception:
        pass  # No running Excel instance found
    return False




def std_codes(code):
    if str(code).replace('.','').isdigit():
        return(str(int(code))).replace('_x000D_\n', '').replace('\n', '')
    else:
        return str(code).replace('_x000D_\n', '').replace('\n', '')


def main(developer=False, nome_relatorio=None):

    if developer:
        ans='S'
    else:
        print('Deseja continuar para criação do relatório? (s/n)')
        ans = str(input())
    
    if ans.upper() != 'S':
        print('Criação de relatório cancelada')
        time.sleep(0.5)
        return 0

    if nome_relatorio is None:
        print('Digite um nome para o novo relatório:')
        nome_relatorio = str(input())

    
    current_script_directory = os.path.dirname(os.path.abspath(__file__))
    main_dir = '/'.join(current_script_directory.split('\\')[:-1]) + '/'
    model_data_folder = main_dir + 'Dados Intermediários/'
    input_folder = main_dir + 'Dados de Input/'
    report_folder = main_dir + 'Relatórios/'

    # 1.2. Configurações gerais
    for file in os.listdir(main_dir):
        if len(file) > 4:
            if file[-5:] == '.xlsm':
                workbook = file
    try:
        close_excel_file_if_open(main_dir + workbook)
        wb = xw.Book(main_dir + workbook)
    except:
        wb = xw.Book(workbook)

    # Lendo frequências
    frequency_sheet = wb.sheets['Frequências']

    freq_df = frequency_sheet.range("B6").expand().value
    freq_df = pd.DataFrame(freq_df[1:], columns=freq_df[0])
    freq_df['FILIAL'] = freq_df['FILIAL'].astype(str)
    freq_df['PARCEIRO'] = freq_df['PARCEIRO'].apply(std_codes)
    freq_df['PATRIMONIO'] = freq_df['PATRIMONIO'].apply(std_codes)
    freq_df['Frequência Mínima (visitas/semana)'] = freq_df['Frequência Mínima (visitas/semana)'].astype(int)
    freq_df['Frequência de Reposição (visitas/semana)'] = freq_df['Frequência de Reposição (visitas/semana)'].astype(int)
    freq_df['Frequência (visitas/semana)'] = freq_df['Frequência (visitas/semana)'].astype(int)
    
    wb_piloto=wb

    # Cria um novo workbook (arquivo Excel)
    wb = xw.Book()  # Cria uma nova pasta de trabalho
    app = xw.apps.active

    ws = wb.sheets[0]
    ws.name = "Consumo"
    ws.api.Activate()  # Ativa a planilha
    app.api.ActiveWindow.DisplayGridlines = False  # Remove as gridline
    ws = wb.sheets.add(name="Patrimônios")
    ws.api.Activate()  # Ativa a planilha
    app.api.ActiveWindow.DisplayGridlines = False  # Remove as gridline
    ws = wb.sheets.add(name="Alocação Sugerida (Patrimônios)")
    ws.api.Activate()  # Ativa a planilha
    app.api.ActiveWindow.DisplayGridlines = False  # Remove as gridline
    ws = wb.sheets.add(name="Alocação Sugerida (Livros)")
    ws.api.Activate()  # Ativa a planilha
    app.api.ActiveWindow.DisplayGridlines = False  # Remove as gridline
    ws = wb.sheets.add(name="Livros")
    ws.api.Activate()  # Ativa a planilha
    app.api.ActiveWindow.DisplayGridlines = False  # Remove as gridline
    ws = wb.sheets.add(name="Lotes")
    ws.api.Activate()  # Ativa a planilha
    app.api.ActiveWindow.DisplayGridlines = False  # Remove as gridline
    ws = wb.sheets.add(name="Visão Geral")
    ws.api.Activate()  # Ativa a planilha
    app.api.ActiveWindow.DisplayGridlines = False  # Remove as gridline

    wb.save(report_folder + f"{nome_relatorio}.xlsx")  # Salva o arquivo no diretório atual
    #wb.close()

    #wb = xw.Book(report_folder + f"{nome_relatorio}.xlsx")

    result_df = pd.read_parquet(model_data_folder + 'result_livros.parquet')
    result_livros_df = pd.read_parquet(model_data_folder + 'result_livros_resumo.parquet')
    alocation_df = pd.read_parquet(model_data_folder + 'alocacao.parquet')
    aloc_patr_df = pd.read_parquet(model_data_folder + 'alocacao_patrimonios.parquet')
    consumo_med_df = pd.read_parquet(model_data_folder + 'consumo_medio.parquet')
    parceiros_df = pd.read_excel(model_data_folder + 'Dados.xlsx', sheet_name='parceiros')
    parceiros_df['PARCEIRO'] = parceiros_df['PARCEIRO'].apply(std_codes)

    rename_dict = {
        'FILIAL':'Filial', 'PERIODO':'Dia', 'LIVRO':'Livro','ABASTECEDOR':'Abastecedor', 'ESCALA':'Escala Requerida', 'MODAL':'Modal de Transporte', 
        'VISITA':'# Visita', 'PARCEIRO':'Parceiro', 'PATRIMONIO':'Patrimônio', 
        'DIST':'Distância (km)', 'TEMPO_DESLOCAMENTO':'Tempo de Deslocamento (min)', 'TEMPO_SERVICO':'Tempo de Serviço (min)', 'LAT':'Latitude', 'LON':'Longitude'
    }

    livros_df = result_df[result_df['VISITA'] != 0].rename(columns=rename_dict)[rename_dict.values()].copy()


    rename_dict = {
        'FILIAL':'Filial', 'PERIODO':'Dia', 'LIVRO':'Livro','ABASTECEDOR':'Abastecedor', 'ESCALA':'Escala Requerida', 'MODAL':'Modal de Transporte', 'HORAS_DIARIAS':'Horas Diárias', 'FTE':'FTE', 'PATRIMONIOS':'# Patrimônios', 
        'DIST':'Distância (km)', 'TEMPO_DESLOCAMENTO':'Tempo de Deslocamento (min)', 'TEMPO_SERVICO':'Tempo de Serviço (min)'
    }

    lotes_df = result_livros_df.rename(columns=rename_dict)[rename_dict.values()].copy()
    lotes_df['FTE'] = lotes_df['FTE'].round(3)


    resumo_df = (
        result_livros_df
        .groupby(['FILIAL'])
        .agg({'PERIODO':'nunique', 'PATRIMONIOS':'sum', 'HORAS_DIARIAS':'sum', 'FTE':'sum', 'DIST':'sum', 'TEMPO_SERVICO':'sum', 'TEMPO_DESLOCAMENTO':'sum'})
        .reset_index()
        .merge(
            result_df.groupby(['FILIAL']).agg({'PATRIMONIO':'nunique'}).reset_index(),
            on='FILIAL',
            how='left'
        )
    )

    for col in ['HORAS_DIARIAS', 'FTE', 'DIST', 'TEMPO_SERVICO', 'TEMPO_DESLOCAMENTO']:
        resumo_df[col] = resumo_df[col]/resumo_df['PERIODO']

    resumo_df['FREQ'] = round(resumo_df['PATRIMONIOS']/resumo_df['PATRIMONIO'],2)

    resumo_df['TEMPO_SERVICO'] = round(resumo_df['TEMPO_DESLOCAMENTO']/60,1)
    resumo_df['TEMPO_SERVICO'] = round(resumo_df['TEMPO_DESLOCAMENTO']/60,1)
    resumo_df['DIST'] = round(resumo_df['DIST'],1)


    rename_dict = {
        'FILIAL':'Filial', 'PATRIMONIO':'# Patrimonios', 'FREQ':'Frequência Média Semanal', 'HORAS_DIARIAS':'HH por Dia', 'HORAS_DIARIAS':'FTEs Requeridos por Dia',
        'TEMPO_DESLOCAMENTO':'Tempo de Deslocamento por Dia (h)','TEMPO_SERVICO':'Tempo de Serviço por Dia (h)','DIST':'Deslocamento por Dia (km)'
    }

    resumo_df = resumo_df.rename(columns=rename_dict)[rename_dict.values()].sort_values(by='# Patrimonios', ascending=False)

    filiais = resumo_df['Filial'].to_list()


    livros_por_dia = (
        result_livros_df
        .groupby(['FILIAL', 'PERIODO'])
        .agg({'LIVRO':'nunique'})
        .reset_index()
        .pivot_table(index=['PERIODO'], columns='FILIAL', values='LIVRO')
        .fillna(0)
        .astype(int)
        .reset_index()
        .rename(columns={'PERIODO':'Dia'})
    )
    livros_por_dia = livros_por_dia[['Dia'] + filiais]


    fte_por_dia = (
        result_livros_df
        .groupby(['FILIAL', 'PERIODO'])
        .agg({'FTE':'sum'})
        .reset_index()
        .pivot_table(index=['PERIODO'], columns='FILIAL', values='FTE')
        .fillna(0)
        .round(1)
        .reset_index()
        .rename(columns={'PERIODO':'Dia'})
    )
    fte_por_dia = fte_por_dia[['Dia'] + filiais]


    escalas_por_dia = (
        result_livros_df
        .groupby(['FILIAL', 'PERIODO', 'ESCALA', 'MODAL'])
        .agg({'LIVRO':'nunique'})
        .reset_index()
        .pivot_table(index=['PERIODO', 'ESCALA', 'MODAL'], columns='FILIAL', values='LIVRO')
        .fillna(0)
        .astype(int)
        .reset_index()
        .rename(columns={'PERIODO':'Dia', 'ESCALA':'Escala', 'MODAL':'Modal'})
    )

    escalas_por_dia = escalas_por_dia[['Dia', 'Escala', 'Modal'] + filiais]


    patrimonios_df = (
        freq_df
        .merge(
            parceiros_df[['FILIAL', 'PARCEIRO', 'INICIO_FUNCIONAMENTO', 'FIM_FUNCIONAMENTO', 'LAT', 'LON']],
            on=['FILIAL', 'PARCEIRO'],
            how='left'
        )
        .rename(columns={'FILIAL':'Filial','PARCEIRO':'Parceiro', 'PATRIMONIO':'Patrimônio','INICIO_FUNCIONAMENTO':'Horário Inicial de Funcionamento', 'FIM_FUNCIONAMENTO':'Horário Final de Funcionamento', 'LAT':'Latitude', 'LON':'Longitude'})
    )

    patrimonios_df['Horário Inicial de Funcionamento'] = patrimonios_df['Horário Inicial de Funcionamento'].astype(str)
    patrimonios_df['Horário Final de Funcionamento'] = patrimonios_df['Horário Final de Funcionamento'].astype(str)


    rename_dict = {
        'FILIAL':'Filial', 'PARCEIRO':'Parceiro', 'PATRIMONIO':'Patrimônio', 'INSUMO':'Insumo', 
        'CONSUMO_SEMANAL':'Consumo Médio Semanal', 'CAPACIDADE':'Capacidade', 'NIVEL_REPOSICAO':'Nível de Reposição'
    }

    consumo_df = consumo_med_df.rename(columns=rename_dict)[rename_dict.values()].copy()


    consumo_df['Giro (semanas)'] = consumo_df['Consumo Médio Semanal']/(consumo_df['Capacidade']*(1 - consumo_df['Nível de Reposição']))

    consumo_df['Consumo Médio Semanal'] = consumo_df['Consumo Médio Semanal'].round(1)
    consumo_df['Giro (semanas)'] = consumo_df['Giro (semanas)'].round(2)


    col_w = 15

    # 5.1. Resultados Gerais
    sheet = wb.sheets['Visão Geral']  # Seleciona a aba específica


    sheet.range("A:Z").clear_contents()
    sheet.range("A:Z").api.ClearFormats()

    start_line = 2

    sheet.range(start_line, 2).value = 'Resultados Gerais'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = resumo_df
    rows = len(resumo_df)
    cols = len(resumo_df.columns)

    sheet.range("B:H").columns.autofit()

    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True

    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)



    # 5.2. Livros pro Dia
    start_line = start_line + rows + 3

    sheet.range(start_line, 2).value = 'Livros por Dia'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = livros_por_dia
    rows = len(livros_por_dia)
    cols = len(livros_por_dia.columns)


    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True
    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)


    # 5.3. FTEs por Dia
    start_line = start_line + rows + 3

    sheet.range(start_line, 2).value = 'FTEs por Dia'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = fte_por_dia
    rows = len(fte_por_dia)
    cols = len(fte_por_dia.columns)


    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True
    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)


    # 5.4. Escalas por Dia
    start_line = start_line + rows + 3

    sheet.range(start_line, 2).value = 'Escalas por Dia'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = escalas_por_dia
    rows = len(escalas_por_dia)
    cols = len(escalas_por_dia.columns)


    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True
    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)



    # 5.5. Lotes
    col_w = 20
    sheet = wb.sheets['Lotes']  # Seleciona a aba específica


    sheet.range("A:Z").clear_contents()
    sheet.range("A:Z").api.ClearFormats()

    start_line = 2

    sheet.range(start_line, 2).value = 'Lotes'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = lotes_df
    rows = len(lotes_df)
    cols = len(lotes_df.columns)

    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True

    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)

    sheet.range("B:G").api.NumberFormat = "@"


    
    # 5.6. Livros
    sheet = wb.sheets['Livros']  # Seleciona a aba específica


    sheet.range("A:Z").clear_contents()
    sheet.range("A:Z").api.ClearFormats()

    start_line = 2

    sheet.range(start_line, 2).value = 'Livros'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = livros_df
    rows = len(livros_df)
    cols = len(livros_df.columns)


    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True

    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)

    sheet.range("B:J").api.NumberFormat = "@"


    # 5.7. Patrimônios
    sheet = wb.sheets['Patrimônios']  # Seleciona a aba específica


    sheet.range("A:Z").clear_contents()
    sheet.range("A:Z").api.ClearFormats()

    start_line = 2

    sheet.range(start_line, 2).value = 'Visão de Patrimônios'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = patrimonios_df
    rows = len(patrimonios_df)
    cols = len(patrimonios_df.columns)

    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True

    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)

    sheet.range("B:D").api.NumberFormat = "@"
    sheet.range("H:I").api.NumberFormat = "hh:mm"


    # 5.8. Consumo
    sheet = wb.sheets['Consumo']  # Seleciona a aba específica


    sheet.range("A:Z").clear_contents()
    sheet.range("A:Z").api.ClearFormats()

    start_line = 2

    sheet.range(start_line, 2).value = 'Visão de Consumo'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = consumo_df
    rows = len(consumo_df)
    cols = len(consumo_df.columns)

    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True

    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)

    sheet.range("B:E").api.NumberFormat = "@"
    sheet.range("H:H").api.NumberFormat = "0%"





    # 5.9. Alocação Livros
    sheet = wb.sheets['Alocação Sugerida (Livros)']  # Seleciona a aba específica


    sheet.range("A:Z").clear_contents()
    sheet.range("A:Z").api.ClearFormats()

    start_line = 2

    sheet.range(start_line, 2).value = 'Alocação Sugerida (Livros)'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = alocation_df
    rows = len(alocation_df)
    cols = len(alocation_df.columns)

    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC
    cells.api.NumberFormat = "@"

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True

    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)

    cells.api.NumberFormat = "@"



    # 5.10. Alocação Patirmônios
    sheet = wb.sheets['Alocação Sugerida (Patrimônios)']  # Seleciona a aba específica


    sheet.range("A:Z").clear_contents()
    sheet.range("A:Z").api.ClearFormats()

    start_line = 2

    sheet.range(start_line, 2).value = 'Alocação Sugerida (Livros)'
    sheet.range(start_line, 2).api.Font.Name = "Arial Narrow"
    sheet.range(start_line, 2).api.Font.Bold = True

    # Transferir o DataFrame para o Excel
    sheet.range(start_line + 1, 2).options(index=False).value = aloc_patr_df
    rows = len(aloc_patr_df)
    cols = len(aloc_patr_df.columns)

    cells = sheet.range((start_line + 1, 2), (start_line + 1 + rows, 1 + cols))
    cells.api.Font.Name = "Arial Narrow"

    # Obter o objeto de bordas
    borders = cells.api.Borders

    # Configurar bordas verticais (grossas e brancas)
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    # Configurar bordas horizontais (finas e pretas)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xEFEEEC
    cells.api.NumberFormat = "@"

    cells = sheet.range((start_line + 1, 2), (start_line + 1, 1 + cols))
    cells.api.Font.Bold = True
    cells.column_width = col_w
    cells.api.WrapText = True

    borders = cells.api.Borders

    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xB8B8B8  # Preto
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xB8B8B8  # Preto
    cells.api.Interior.Color = 0xE6E6E6
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)

    cells.api.NumberFormat = "@"



    wb.save()
    
    print('Criação de relatório encerrada')
    
    if developer:
        wb.close()
        wb_piloto.close()
    else:
        input()