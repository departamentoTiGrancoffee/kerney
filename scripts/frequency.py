# ======================================================================================
# Cálculo de frequência de abastecimento (visitas/semana) por patrimônio (máquina)
# --------------------------------------------------------------------------------------
# Fluxo em alto nível:
# 1) Utilidades: fechar Excel aberto; padronizar códigos; ler configs (txt) -> dict
# 2) Função de "repasses": quando a frequência baseada em consumo excede muito a agenda,
#    divide patrimônio/parceiro em _A/_B para permitir janelas distintas no mesmo dia
# 3) main():
#    3.1) Confirmação do usuário (sobrescreve aba "Frequências" no Excel)
#    3.2) Resolução de diretórios/arquivos e abertura do workbook .xlsm
#    3.3) Leitura dos parâmetros na planilha e do arquivo de formato do consumo (txt)
#    3.4) Leitura e padronização de dados (parceiros, patrimônios, insumos, consumo)
#    3.5) Cálculo do consumo médio semanal por (FILIAL, PARCEIRO, PATRIMONIO, INSUMO)
#    3.6) Cálculo de FREQ_BASEADO_EM_CONSUMO (nível de reposição global ou por SKU)
#    3.7) (Opcional) Repasses e ajustes de janelas e depara_point_id
#    3.8) Definição da FREQUENCIA_REPOSICAO respeitando limites (dias/semana, atual)
#    3.9) Aplicação de FREQUENCIA_SEMANAL_MINIMA e flexibilidade (freq_flex)
#    3.10) (Opcional) Padronização intra-parceiro: mesma frequência para todos os ativos
#    3.11) Escrita de saídas (parceiros/patrimônios atualizados, parquet e Excel)
#    3.12) Estilização na aba "Frequências" e finalização
# ======================================================================================

import pandas as pd
import os
import time
from tqdm import tqdm
import xlwings as xw
import numpy as np
import win32com.client

def close_excel_file_if_open(filename):
    # Fecha um arquivo Excel específico se estiver aberto (usa automação COM via pywin32)
    # Evita erro de gravação ao sobrescrever planilhas/arquivos usados pelo Excel.
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
    # Padroniza códigos vindos de Excel/CSV: remove quebras de linha e _x000D_, cast numérico
    if str(code).replace('.','').isdigit():
        return(str(int(code))).replace('_x000D_\n', '').replace('\n', '')
    else:
        return str(code).replace('_x000D_\n', '').replace('\n', '')


def read_txt_as_dict(file_path):
    # Lê um arquivo .txt no formato "chave: valor" e retorna dict com pares limpos.
    """
    Lê um arquivo .txt com linhas no formato `chave: valor` e retorna um dicionário.

    Args:
        file_path (str): Caminho para o arquivo .txt.

    Returns:
        dict: Dicionário com as chaves e valores do arquivo.
    """
    config_dict = {}
    with open(file_path, 'r') as file:
        for line in file:
            line = line.strip()
            if line:  # Ignorar linhas vazias
                key, value = line.split(":", 1)  # Separar chave e valor
                config_dict[key.strip()] = value.strip().strip('"')  # Remover espaços e aspas
                
    return config_dict


def repasses(visitas_df, patrimonios_df, parceiros_df, depara_point_id,
             allowed_patrimonios=None, gap_min_hours=3):
    """
    Divide patrimônios/parceiros em _A/_B quando a FREQ_BASEADO_EM_CONSUMO excede
    significativamente o número de DIAS_POR_SEMANA (fator 1.5), e cria janelas de
    funcionamento com um gap mínimo entre _A e _B.

    Parâmetros
    ----------
    visitas_df : pd.DataFrame
        Base por item/patrimônio com colunas:
        ['FILIAL','PARCEIRO','PATRIMONIO','DIAS_POR_SEMANA','FREQ_BASEADO_EM_CONSUMO', ...]
    patrimonios_df : pd.DataFrame
        Base de patrimônios (aba 'patrimonios').
    parceiros_df : pd.DataFrame
        Base de parceiros (aba 'parceiros') com INICIO_FUNCIONAMENTO e FIM_FUNCIONAMENTO (datetime64[ns]).
    depara_point_id : pd.DataFrame
        De-para geográfico com ['FILIAL','PARCEIRO','POINT_ID','LAT','LON'].
    allowed_patrimonios : list[str] | None
        Lista de patrimônios autorizados a receber repasse (S/N vindo do Excel).
        Se None, considera todos.
    gap_min_hours : int | float
        Gap mínimo entre janelas _A e _B (em horas). Default: 3.

    Retorna
    -------
    (visitas_df, patrimonios_df, parceiros_df, depara_point_id) : tuple[pd.DataFrame]
    """
    # 1) Seleção dos casos que sofrerão repasse
    idx_repasse = (visitas_df['FREQ_BASEADO_EM_CONSUMO'] >
                   visitas_df['DIAS_POR_SEMANA'] * 1.5)  # fator 1.5 mantido

    if allowed_patrimonios is not None:
        idx_repasse = idx_repasse & visitas_df['PATRIMONIO'].isin(allowed_patrimonios)

    # Se nada atende aos critérios, retorna sem alterações
    if not np.any(idx_repasse):
        return visitas_df, patrimonios_df, parceiros_df, depara_point_id

    patrimonios_repasse = visitas_df.loc[idx_repasse, 'PATRIMONIO'].unique().tolist()
    parceiros_repasse = visitas_df.loc[idx_repasse, 'PARCEIRO'].unique().tolist()

    # 2) VISITAS: duplica linhas e reparte consumo em _A e _B
    visitas_df_aux = visitas_df.loc[idx_repasse].copy()
    # _A: trava consumo à capacidade semanal (DIAS_POR_SEMANA)
    visitas_df.loc[idx_repasse, 'PARCEIRO'] = visitas_df.loc[idx_repasse, 'PARCEIRO'] + '_A'
    visitas_df.loc[idx_repasse, 'PATRIMONIO'] = visitas_df.loc[idx_repasse, 'PATRIMONIO'] + '_A'
    visitas_df.loc[idx_repasse, 'FREQ_BASEADO_EM_CONSUMO'] = visitas_df.loc[idx_repasse, 'DIAS_POR_SEMANA']
    # _B: restante
    visitas_df_aux['PARCEIRO'] = visitas_df_aux['PARCEIRO'] + '_B'
    visitas_df_aux['PATRIMONIO'] = visitas_df_aux['PATRIMONIO'] + '_B'
    visitas_df_aux['FREQ_BASEADO_EM_CONSUMO'] = (
        visitas_df_aux['FREQ_BASEADO_EM_CONSUMO'] - visitas_df_aux['DIAS_POR_SEMANA']
    )
    visitas_df = pd.concat([visitas_df, visitas_df_aux], ignore_index=True)

    # 3) PATRIMÔNIOS: ajusta piso mínimo e duplica linhas
    idx_pat = patrimonios_df['PATRIMONIO'].isin(patrimonios_repasse)
    # Piso mínimo pela metade (arredonda para cima)
    patrimonios_df.loc[idx_pat, 'FREQUENCIA_SEMANAL_MINIMA'] = np.ceil(
        patrimonios_df.loc[idx_pat, 'FREQUENCIA_SEMANAL_MINIMA'] / 2
    )

    patrimonios_df_aux = patrimonios_df.loc[idx_pat].copy()

    # _A
    patrimonios_df.loc[idx_pat, 'PATRIMONIO'] = patrimonios_df.loc[idx_pat, 'PATRIMONIO'] + '_A'
    patrimonios_df.loc[idx_pat, 'PARCEIRO'] = patrimonios_df.loc[idx_pat, 'PARCEIRO'] + '_A'
    # _B
    patrimonios_df_aux['PATRIMONIO'] = patrimonios_df_aux['PATRIMONIO'] + '_B'
    patrimonios_df_aux['PARCEIRO'] = patrimonios_df_aux['PARCEIRO'] + '_B'

    # Mantém flags de repasse nos duplicados, se existirem
    if 'APLICA_REPASSE' in patrimonios_df.columns:
        patrimonios_df.loc[idx_pat, 'APLICA_REPASSE'] = 'S'
        patrimonios_df_aux['APLICA_REPASSE'] = 'S'
    if 'APLICA_REPASSE_BOOL' in patrimonios_df.columns:
        patrimonios_df.loc[idx_pat, 'APLICA_REPASSE_BOOL'] = True
        patrimonios_df_aux['APLICA_REPASSE_BOOL'] = True

    patrimonios_df = pd.concat([patrimonios_df, patrimonios_df_aux], ignore_index=True)

    # Lista final de parceiros após duplicação (para filtrar tabelas derivadas)
    parceiros_finais = patrimonios_df['PARCEIRO'].unique().tolist()

    # 4) PARCEIROS: divide janelas com gap mínimo entre _A e _B
    # Seleciona somente os parceiros-base (sem sufixo) que sofrerão repasse
    idx_parc_rep = parceiros_df['PARCEIRO'].isin(parceiros_repasse)
    parceiros_base = parceiros_df.loc[idx_parc_rep].copy()

    # Séries de horários (datetime64[ns])
    start_s = parceiros_base['INICIO_FUNCIONAMENTO']
    end_s = parceiros_base['FIM_FUNCIONAMENTO']

    dur_s = (end_s - start_s)
    gap_des = pd.to_timedelta(gap_min_hours, unit='h')

    # Ajusta gap quando janela total é curta (mantém viabilidade com janelas não vazias)
    gap_aplicado = pd.Series(gap_des, index=parceiros_base.index)
    impossiveis = dur_s <= (gap_des + pd.to_timedelta(1, unit='m'))
    if impossiveis.any():
        print("[repasses] Aviso: alguns parceiros não têm duração suficiente para o gap mínimo; reduzindo gap nesses casos.")
        gap_aplicado.loc[impossiveis] = np.maximum(
            pd.to_timedelta(0, unit='m'),
            dur_s.loc[impossiveis] - pd.to_timedelta(1, unit='m')
        )

    # Ponto médio “efetivo” para dividir o intervalo restante (dur - gap) ao meio
    mid_s = start_s + (dur_s - gap_aplicado) / 2

    # Monta _A (manhã) e _B (tarde) com o gap no meio
    parceiros_A = parceiros_base.copy()
    parceiros_B = parceiros_base.copy()

    parceiros_A['PARCEIRO'] = parceiros_A['PARCEIRO'] + '_A'
    parceiros_B['PARCEIRO'] = parceiros_B['PARCEIRO'] + '_B'

    parceiros_A['INICIO_FUNCIONAMENTO'] = start_s
    parceiros_A['FIM_FUNCIONAMENTO'] = mid_s - gap_aplicado

    parceiros_B['INICIO_FUNCIONAMENTO'] = mid_s + gap_aplicado
    parceiros_B['FIM_FUNCIONAMENTO'] = end_s

    # Reconstroi parceiros_df: (sem os base) + (A e B), e filtra para os parceiros finais
    parceiros_sem_base = parceiros_df.loc[~idx_parc_rep].copy()
    parceiros_df = pd.concat([parceiros_sem_base, parceiros_A, parceiros_B], ignore_index=True)
    parceiros_df = parceiros_df.loc[parceiros_df['PARCEIRO'].isin(parceiros_finais)].reset_index(drop=True)

    # 5) DEPARA_POINT_ID: duplica mapeamentos para _A/_B e filtra para parceiros finais
    idx_depara = depara_point_id['PARCEIRO'].isin(parceiros_repasse)
    depara_A = depara_point_id.loc[idx_depara].copy()
    depara_B = depara_point_id.loc[idx_depara].copy()

    depara_A['PARCEIRO'] = depara_A['PARCEIRO'] + '_A'
    depara_B['PARCEIRO'] = depara_B['PARCEIRO'] + '_B'

    depara_point_id = pd.concat([depara_point_id, depara_A, depara_B], ignore_index=True)
    depara_point_id = depara_point_id.loc[depara_point_id['PARCEIRO'].isin(parceiros_finais)].reset_index(drop=True)

    return visitas_df, patrimonios_df, parceiros_df, depara_point_id

def main(developer=False, visitas_duplas=False, nivel_reposicao=None,
         padronizacao_intra_patrimonio=False, freq_flex=None):
    # ----------------------------------------------------------------------------------
    # Parâmetros:
    #   developer: pula interação de confirmação
    #   visitas_duplas: ativa lógica de repasses (divisão _A/_B)
    #   nivel_reposicao: se fornecido, usa NÍVEL DE REPOSIÇÃO global (0..1) p/ todos SKUs
    #   padronizacao_intra_patrimonio: força mesma frequência entre patrimônios do mesmo parceiro
    #   freq_flex: flexibilidade (+/-) em relação à frequência atual mínima por patrimônio
    # Saídas: escreve resultados na aba "Frequências" e arquivos auxiliares em "Dados Intermediários"
    # ----------------------------------------------------------------------------------
    
    if developer:
        ans='S'
    else:
        print('Essa Operação Sobrescreverá os dados da aba "Frequências", deseja continuar? (s/n)')
        ans = str(input())
    
    if ans.upper() != 'S':
        # Cancelamento controlado pelo usuário
        print('Calculo de Frequência Cancelado')
        time.sleep(0.5)
        return 0

    start = time.time()
    print("Iniciando Cálculo de Frequência")

    section_start = time.time()
    print("Lendo dados de Input e Configurações...")

    # 1) Resolução de diretórios e localização do workbook .xlsm
    current_script_directory = os.path.dirname(os.path.abspath(__file__))
    main_dir = '/'.join(current_script_directory.split('\\')[:-1]) + '/'
    model_data_folder = main_dir + 'Dados Intermediários/'
    input_folder = main_dir + 'Dados de Input/'

    # Descobrir arquivo .xlsm na pasta raiz do projeto (main_dir)
    for file in os.listdir(main_dir):
        if len(file) > 4:
            if file[-5:] == '.xlsm':
                workbook = file
                
    # Garante que o Excel não está com o arquivo aberto (evita lock ao gravar)
    close_excel_file_if_open(main_dir + workbook)
    wb = xw.Book(main_dir + workbook)
    config_sheet = wb.sheets['Configurações']

    # 2) Ler parâmetros de configuração da planilha
    dias_semana = int(config_sheet['D7'].value)
    input_mapping = {"Sim": True, "Não": False}
    visitas_duplas = input_mapping[config_sheet['J10'].value]

    # 3) Ler arquivo de configuração de formato do CSV de consumo
    read_config = read_txt_as_dict(input_folder + 'formato_consumo.txt')

    # 4) Importar dados de entrada (com padronização de códigos)
    parceiros_df = pd.read_excel(input_folder + 'Dados.xlsx', sheet_name='parceiros')
    parceiros_df['PARCEIRO'] = parceiros_df['PARCEIRO'].apply(std_codes)
    parceiros_df['INICIO_FUNCIONAMENTO'] = pd.to_datetime(parceiros_df['INICIO_FUNCIONAMENTO'], format='%H:%M:%S')
    parceiros_df['FIM_FUNCIONAMENTO'] = pd.to_datetime(parceiros_df['FIM_FUNCIONAMENTO'], format='%H:%M:%S')

    patrimonios_df = pd.read_excel(input_folder + 'Dados.xlsx', sheet_name='patrimonios')
    patrimonios_df['PATRIMONIO'] = patrimonios_df['PATRIMONIO'].apply(std_codes)
    patrimonios_df['PARCEIRO'] = patrimonios_df['PARCEIRO'].apply(std_codes)

    # --- NOVO: coluna APLICA_REPASSE (S/N) + validação + bool auxiliar ---
    if 'APLICA_REPASSE' not in patrimonios_df.columns:
        patrimonios_df['APLICA_REPASSE'] = 'N'   # default quando a coluna não existe
    else:
        patrimonios_df['APLICA_REPASSE'] = (
            patrimonios_df['APLICA_REPASSE']
            .astype(str)
            .str.strip()
            .str.upper()
        )
        invalid = ~patrimonios_df['APLICA_REPASSE'].isin(['S','N'])
        if invalid.any():
            linhas_invalidas = patrimonios_df.loc[invalid, ['FILIAL','PARCEIRO','PATRIMONIO','APLICA_REPASSE']]
            raise ValueError(
                "Coluna APLICA_REPASSE deve conter apenas 'S' ou 'N'. "
                f"Exemplos inválidos:\n{linhas_invalidas.head(10).to_string(index=False)}"
            )
    patrimonios_df['APLICA_REPASSE_BOOL'] = (patrimonios_df['APLICA_REPASSE'] == 'S')

    insumos_df = pd.read_excel(input_folder + 'Dados.xlsx', sheet_name='insumos')
    insumos_df['PATRIMONIO'] = insumos_df['PATRIMONIO'].apply(std_codes)
    insumos_df['PARCEIRO'] = insumos_df['PARCEIRO'].apply(std_codes)
    insumos_df['INSUMO'] = insumos_df['INSUMO'].apply(std_codes)

    consumo_df = pd.read_csv(input_folder + 'Dados de Consumo.csv', encoding=read_config['encoding'], sep=read_config['sep'], decimal=read_config['decimal']).dropna()
    consumo_df['PATRIMONIO'] = consumo_df['PATRIMONIO'].apply(std_codes)
    consumo_df['PARCEIRO'] = consumo_df['PARCEIRO'].apply(std_codes)
    consumo_df['INSUMO'] = consumo_df['INSUMO'].apply(std_codes)

    # De-para geográfico gerado na etapa de matriz de distâncias
    depara_point_id = pd.read_parquet(model_data_folder + 'depara_point_id.parquet')[['FILIAL', 'PARCEIRO', 'POINT_ID', 'LAT', 'LON']]
    depara_point_id['PARCEIRO'] = depara_point_id['PARCEIRO'].apply(std_codes)

    print('Leitura de dados encerrada, tempo: {:.1f}s'.format(time.time() - section_start))
    section_start = time.time()
    print('Iniciando cálculo de frequência...')

    # 5) Preparação do consumo médio semanal por patrimônio e insumo
    consumo_med_df = consumo_df.copy()
    for col in ['INICIO', 'FIM']:
        # Se datas vierem como string, converter dd/mm/YYYY
        if str(consumo_med_df[col].dtypes) == 'object':
            consumo_med_df[col] = pd.to_datetime(consumo_med_df[col], format='%d/%m/%Y')
    consumo_med_df['INICIO'] = consumo_med_df['INICIO'].astype('datetime64[ns]')
    consumo_med_df['FIM'] = consumo_med_df['FIM'].astype('datetime64[ns]')
    consumo_med_df['DIAS_TOTAL'] = np.maximum(1, (consumo_med_df['FIM'] - consumo_med_df['INICIO']).dt.days)
    consumo_med_df['SEMANAS'] = consumo_med_df['DIAS_TOTAL']/7
    consumo_med_df = consumo_med_df.groupby(['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'INSUMO']).agg({'CONSUMO':'sum', 'SEMANAS':'sum'}).reset_index()
    consumo_med_df['CONSUMO_SEMANAL'] = consumo_med_df['CONSUMO']/consumo_med_df['SEMANAS']
    consumo_med_df = (consumo_med_df
                    .merge(insumos_df,
                            on=['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'INSUMO'],
                            how='left')
                    .merge(patrimonios_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'DIAS_POR_SEMANA', 'FREQUENCIA_ATUAL']],
                            on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
                            how='outer')
                    .loc[lambda x: x['DIAS_POR_SEMANA'].notna()]
                    .reset_index(drop=True))
    
    # 6) Cálculo base por item (cada linha é um SKU/insumo do patrimônio)
    visitas_df = consumo_med_df.copy()

    # Valores default para preenchimento de faltas (observa-se 'INSUMO ' com espaço à direita)
    fill_values = {'INSUMO': "COPO", 'CONSUMO': 0, 'SEMANAS':1 ,'CAPACIDADE':1, 'MAQUINAS':"-", 'NIVEL_REPOSICAO':0 , 'CONSUMO_SEMANAL': 0} # Aplicando fillna com dicionário
    visitas_df = visitas_df[visitas_df["CAPACIDADE"].notna()]
    visitas_df = visitas_df.fillna(fill_values)
 
    # 7) Dois modos de calcular FREQ_BASEADO_EM_CONSUMO:
    #    a) Se nivel_reposicao (global) vier: usa esse valor para todos os itens
    #    b) Caso contrário: usa NIVEL_REPOSICAO por linha (insumo)
    if nivel_reposicao is not None:
        # a) Usando nível de reposição global (0..1)
        visitas_df['FREQ_BASEADO_EM_CONSUMO'] = np.ceil(visitas_df['CONSUMO_SEMANAL']/(visitas_df['CAPACIDADE']*(1 - nivel_reposicao)))
        # Consolidar por patrimônio: usar a MAIOR frequência entre insumos (max)
        visitas_df = visitas_df.groupby(['FILIAL', 'PARCEIRO', 'PATRIMONIO']).agg({'FREQ_BASEADO_EM_CONSUMO':'max'}).reset_index()
        # Trazer FREQUENCIA_SEMANAL_MINIMA para aplicar piso
        visitas_df = (patrimonios_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'FREQUENCIA_SEMANAL_MINIMA']]
            .merge(
                visitas_df,
                on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
                how='left'))
        # Preencher faltas com o mínimo semanal
        visitas_df['FREQ_BASEADO_EM_CONSUMO'] = (
            visitas_df['FREQ_BASEADO_EM_CONSUMO']
            .fillna(visitas_df['FREQUENCIA_SEMANAL_MINIMA']))
        wb.close()
        return visitas_df
    else:
        # b) Usando NIVEL_REPOSICAO específico de cada linha
        visitas_df['FREQ_BASEADO_EM_CONSUMO'] = np.ceil(visitas_df['CONSUMO_SEMANAL']/(visitas_df['CAPACIDADE']*(1 - visitas_df['NIVEL_REPOSICAO'])))
    

    # 8) (Opcional) Repasses: divide patrimônio e janelas do parceiro (_A/_B) quando habilitado
    if visitas_duplas:
        # apenas patrimônios marcados como 'S' entram no repasse
        patrimonios_autorizados = (
            patrimonios_df.loc[patrimonios_df['APLICA_REPASSE_BOOL'], 'PATRIMONIO']
            .unique().tolist()
        )
        visitas_df, patrimonios_df, parceiros_df, depara_point_id = repasses(
            visitas_df, patrimonios_df, parceiros_df, depara_point_id,
            allowed_patrimonios=patrimonios_autorizados
        )

    # 9) FREQUENCIA_REPOSICAO é limitada por DIAS_POR_SEMANA e pela FREQUENCIA_ATUAL
    #    (evita propor mais visitas que a janela semanal e respeita estado atual)
    visitas_df['FREQUENCIA_REPOSICAO'] = np.minimum(visitas_df['FREQ_BASEADO_EM_CONSUMO'],
                                                    np.minimum(visitas_df['DIAS_POR_SEMANA'],
                                                               visitas_df['FREQUENCIA_ATUAL'])).astype(int)

    # Consolidar por patrimônio pegando o máximo entre insumos (linha -> patrimônio)
    visitas_df = visitas_df.groupby(['FILIAL', 'PARCEIRO', 'PATRIMONIO']).agg({'FREQUENCIA_REPOSICAO':'max'}).reset_index()

    # 10) Aplicar FREQUENCIA_SEMANAL_MINIMA (piso regulatório/operacional por patrimônio)
    visitas_df = (patrimonios_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'FREQUENCIA_SEMANAL_MINIMA', 'FREQUENCIA_ATUAL']]
                .merge(
                    visitas_df,
                    on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
                    how='left')
                .fillna(0)
                )
    
    # 11) Se houver flexibilidade (freq_flex), ajusta o piso mínimo com base na atual
    if freq_flex is not None:
        visitas_df['FREQUENCIA_SEMANAL_MINIMA'] = np.maximum(visitas_df['FREQUENCIA_SEMANAL_MINIMA'], (visitas_df['FREQUENCIA_ATUAL']-freq_flex))
    
    # 12) FREQUENCIA final = max(piso mínimo, frequência de reposição calculada)
    visitas_df['FREQUENCIA'] = np.maximum(visitas_df['FREQUENCIA_SEMANAL_MINIMA'], visitas_df['FREQUENCIA_REPOSICAO'])
    
    # 13) (Opcional) Padronização intra-parceiro: força mesma frequência para todos patrimônios
    if padronizacao_intra_patrimonio:
        visitas_df_aux = (
            visitas_df.groupby('PARCEIRO')['FREQUENCIA'].max()
            .reset_index().rename(columns={'FREQUENCIA':'FREQUENCIA_PARC'}))
        visitas_df = (visitas_df.merge(visitas_df_aux, on='PARCEIRO', how='left')
                      .assign(FREQUENCIA= lambda x: x['FREQUENCIA_PARC'])
                      .drop(columns='FREQUENCIA_PARC'))
        visitas_df["FREQUENCIA"]=visitas_df["FREQUENCIA"].astype(int)
    
    # 14) Renomear colunas para nomenclatura de apresentação (mantém subset de interesse)
    renames = {'FILIAL':'FILIAL', 'PARCEIRO': 'PARCEIRO', 'PATRIMONIO': 'PATRIMONIO', 
               'FREQUENCIA_ATUAL': 'Frequência Atual (visitas/semana)', 'FREQUENCIA_SEMANAL_MINIMA': 'Frequência Mínima (visitas/semana)',
               'FREQUENCIA_REPOSICAO': 'Frequência de Reposição (visitas/semana)', 'FREQUENCIA': 'Frequência (visitas/semana)'}
    visitas_df = visitas_df[renames.keys()].rename(columns=renames)

    # 15) Converter colunas de horário para time puro (remove data 1900-01-01)
    parceiros_df['INICIO_FUNCIONAMENTO'] = (parceiros_df['INICIO_FUNCIONAMENTO'].dt.time) #.astype(str)
    parceiros_df['FIM_FUNCIONAMENTO'] = (parceiros_df['FIM_FUNCIONAMENTO'].dt.time) #.astype(str)

    # 16) Persistir ajustes intermediários:
    #     - Grava parceiros/patrimônios atualizados em "Dados Intermediários/Dados.xlsx"
    #     - Atualiza depara_point_id (caso repasses)
    with pd.ExcelWriter(model_data_folder+'Dados.xlsx', engine="xlsxwriter") as writer:
        parceiros_df.to_excel(writer, sheet_name="parceiros", index=False)
        patrimonios_df.to_excel(writer, sheet_name="patrimonios", index=False)

    depara_point_id.to_parquet(model_data_folder+'depara_point_id_atualizado.parquet', index=False)

    print('Cálculo de frequência encerrado, tempo: {:.1f}s'.format(time.time() - section_start))
    section_start = time.time()
    print('Salvando dados e escrevendo resultados na aba "Frequências"...')

    # 17) Escrever resultado na aba "Frequências" do workbook e aplicar formatação
    sheet = wb.sheets['Frequências']  # Seleciona a aba específica

    # Transferir o DataFrame para o Excel
    sheet.range("7:1048576").clear_contents()
    sheet['B6'].options(index=False).value = visitas_df  # Transfere o DataFrame com o cabeçalho

    # Selecionar o intervalo de células a ser formatado (borda e estilo)
    cells = sheet.range('B7:H{}'.format(len(visitas_df) + 6))

    # Bordas (usa API COM do Excel)
    borders = cells.api.Borders

    # Bordas laterais e horizontais (cores em BGR)
    borders(11).Weight = 4  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco

    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xD2D2D2  # Preto
    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xD2D2D2  # Preto
    borders(9).Weight = 4  # xlEdgeBottom, fina
    borders(9).Color = 0xD2D2D2  # Preto

    # Destaque visual na coluna H (Frequência final)
    cells = sheet.range('H7:H{}'.format(len(visitas_df) + 6))
    cells.api.Font.Bold = True                # Negrito
    cells.api.Font.Italic = True
    cells.api.Font.Size = 12                  # Tamanho da fonte
    cells.api.Interior.Color = 0xF4EDFC

    # 18) Persistir série temporal agregada (consumo médio) para reuso posterior
    consumo_med_df.to_parquet(model_data_folder + 'consumo_medio.parquet', index=False)

    print('Dados salvos com sucesso, tempo: {:.1f}s'.format(time.time() - section_start))

    # 19) Encerramento
    print('Cálculo de Frequências Encerrado, tempo total: {:.1f}s\nPrecione Enter para continuar...'.format(time.time() - start))
    
    # Fecha workbook ou aguarda ENTER conforme modo
    if developer:
        wb.close()
    else:
        input()
