# ======================================================================================
# Cálculo de matrizes de distância/tempo (carro e a pé) entre clientes/parceiros
# --------------------------------------------------------------------------------------
# Fluxo:
# 1) Ler inputs, padronizar códigos e gerar IDs de pontos (POINT_ID)
# 2) Montar uma malha euclidiana limitada (até max_dist e top max_viz vizinhos)
# 3) Planejar chamadas à ORS Distance Matrix respeitando limite de payload (api_limit)
# 4) Executar chamadas (com rate limit) e persistir resultados brutos por modal
# 5) Tratar dados faltantes estimando DIST (aprox. planar) e preenchendo DISTANCE/DURATION
# 6) Salvar matrizes finais por modal (carro/a_pe)
# Observações importantes estão marcadas com "ATENÇÃO".
# ======================================================================================

import pandas as pd
import os
import math
import json
import time
from tqdm import tqdm
import openrouteservice
import numpy as np


def std_codes(code):
    """
    Padroniza códigos vindos de Excel (remove quebras de linha e _x000D_).
    Se for número (inclui floats em string), converte para inteiro sem casas.
    """
    if str(code).replace('.','').isdigit():
        return(str(int(code))).replace('_x000D_\n', '').replace('\n', '')
    else:
        return str(code).replace('_x000D_\n', '').replace('\n', '')


def main(developer=False):
    """
    Executa o pipeline completo para os dois perfis:
      - 'carro'  -> 'driving-car'
      - 'a_pe'   -> 'foot-walking'
    Para cada perfil (modal), gera/retoma chamadas, consulta a ORS e trata a matriz.
    """

    # 0) Mapeia rótulos internos -> perfis da ORS
    profile_dict={"carro": "driving-car",
                  "a_pe": "foot-walking"}
    
    # Loop por modal (carro e a pé)
    for modal in profile_dict.keys():
        
        # 1) Paths baseados na localização do script
        current_script_directory = os.path.dirname(os.path.abspath(__file__))
        main_dir = '/'.join(current_script_directory.split('\\')[:-1])
        model_data_folder = main_dir + '/Dados Intermediários/'
        input_folder = main_dir + '/Dados de Input/'

        # Controle de retomada de execução
        ans = 2

        # 2) Interação com usuário para retomar/iniciar/cancelar (quando não estiver em modo developer)
        if not developer:
            if 'distance_matrix_raw.parquet' in os.listdir(model_data_folder):
                print('Um conjunto de chamadas já foi realizado anteriormente, deseja retomar o processo anterior?')
                print('Digite 1 para retomar o processo anterior')
                print('Digite 2 para iniciar um novo processo')
                print('Digite 3 para cancelar a operação')
                ans = int(input())
                if (ans != 1) and (ans != 2):
                    print('Operação Cancelada')
                    return 0
            else:
                print('Iniciando Chamadas de API, deseja continuar? (S/N)')
                ans_2 = int(input())
                if ans_2.upper() != 'S':
                    print('Operação Cancelada')
                    return 0

        # 3) Preparação de lista de chamadas e pré-processamentos
        if ans == 2:
            print('Gerando lista de chamadas para API')

            # Parâmetros-chave:
            max_dist = 15        # raio de corte (km) 
            max_viz = 200        # máximo de vizinhos (top-N) por origem
            api_limit = 3500     # limite de carga por chamada (produto |sources| x |destinations| aproximado)

            # 3.1) Ler parceiros (clientes) e padronizar código/ID
            input_parceiros_df = pd.read_excel(input_folder + 'Dados.xlsx', sheet_name='parceiros')
            input_parceiros_df['PARCEIRO'] = input_parceiros_df['PARCEIRO'].apply(std_codes)

            # Observação: blocos abaixo foram deixados comentados como no código original.
            # Poderiam incluir residências dos abastecedores como pontos adicionais.
            # input_residencias_df = pd.read_excel(input_folder + 'Dados.xlsx', sheet_name='residencias')
            # input_residencias_df['PARCEIRO'] = input_residencias_df['ABASTECEDOR'].apply(std_codes)
            # input_residencias_df = input_residencias_df[['FILIAL', 'PARCEIRO', 'LAT', 'LON']].copy()
            # input_residencias_df = input_residencias_df[((input_residencias_df['LAT'].notna()) &
            #                                              (input_residencias_df['LON'].notna()))].copy()
            # input_parceiros_df = pd.concat([input_parceiros_df, input_residencias_df]).reset_index(drop=True)

            # 3.2) Carregar chave da ORS (mantida leitura de arquivo local)
            with open(input_folder + "ORS API KEY.txt", "r") as file:
                API_KEY = file.read()

            # 3.3) Criar de-para de POINT_ID (hash geo simplificado com sinal e 4 casas decimais em milésimos de grau)
            #      Formato: N/E/W/S + valor absoluto * 10000 com zero-fill -> garante ID determinístico por coord.
            depara_point_id = input_parceiros_df.copy()
            depara_point_id['POINT_ID'] = (
                depara_point_id['LAT'].apply(lambda x: {True:'N',False:'S'}[(x>=0)] + str(int(abs(x)*10000)).zfill(6))
                + depara_point_id['LON'].apply(lambda x: {True:'E',False:'W'}[(x>=0)] + str(int(abs(x)*10000)).zfill(6))
                # + depara_point_id['PARCEIRO']  # Mantido comentado conforme original
            )

            # 3.4) Consolidar pontos únicos por filial
            points_df = depara_point_id[['FILIAL','POINT_ID', 'LAT', 'LON']].drop_duplicates(subset=['POINT_ID','FILIAL']).copy()

            # 3.5) Persistir de-para para uso posterior
            depara_point_id.to_parquet(model_data_folder + 'depara_point_id.parquet', index=False)

            # 3.6) Preparar malha euclidiana (pré-filtro de pares I–J próximos) para reduzir chamadas à API
            euc_matrix_df = (
                depara_point_id.drop_duplicates(subset=['FILIAL','POINT_ID'])[['FILIAL','POINT_ID','LAT','LON']].copy()
                .merge(
                    depara_point_id.drop_duplicates(subset=['FILIAL','POINT_ID'])[['FILIAL','POINT_ID','LAT','LON']].copy(),
                    on=['FILIAL'],
                    how='left',
                    suffixes=['_I', '_J']
                )
            )

            # Conversão de 1 grau latitudinal em km (aprox. 40.000 km / 360)
            lat_to_km = 40000/360

            # 3.7) Distância euclidiana aproximada em km (ajuste longitudinal por cos(lat média))
            euc_matrix_df['DIST'] = (
                ((euc_matrix_df['LAT_I'] - euc_matrix_df['LAT_J'])*lat_to_km)**2 + 
                ((euc_matrix_df['LON_I'] - euc_matrix_df['LON_J'])*
                    (lat_to_km*(((math.pi/180) * (euc_matrix_df['LAT_I'] + euc_matrix_df['LAT_J'])/2).apply(math.cos))))**2
            )**(1/2)
            
            # 3.8) Cortar por raio e manter top-N vizinhos por origem
            euc_matrix_df = euc_matrix_df[(euc_matrix_df['DIST'] <= max_dist)].sort_values(by=['FILIAL','POINT_ID_I', 'DIST'])
            euc_matrix_df['COUNT'] = 1
            euc_matrix_df['COUNT'] = euc_matrix_df.groupby(['FILIAL', 'POINT_ID_I'])['COUNT'].cumsum()
            euc_matrix_df = euc_matrix_df[euc_matrix_df['COUNT'] <= max_viz].copy()

            # 3.9) Montar calls por filial, agregando sources/destinations até api_limit
            calls = {}
            filiais = euc_matrix_df['FILIAL'].unique().tolist()
            for filial in filiais:

                filial_df = euc_matrix_df[euc_matrix_df['FILIAL'] == filial][['POINT_ID_I', 'POINT_ID_J', 'DIST']].copy()

                point_list = filial_df['POINT_ID_I'].unique().tolist()

                point_i = filial_df['POINT_ID_I'].to_list()
                point_j = filial_df['POINT_ID_J'].to_list()

                # Mapa de vizinhos próximos por origem
                close_points = {i:[] for i in point_list}
                for i in range(len(point_i)):
                    close_points[point_i[i]].append(point_j[i])

                # Inicialização da lista de pacotes de chamadas
                filial_calls = []
                unsigned_sources = list(point_list)

                # Primeira chamada com o primeiro source
                filial_calls =[{
                            'sources':[unsigned_sources[0]],
                            'destinations':list(close_points[unsigned_sources[0]])
                        }]

                last_source = unsigned_sources[0]
                unsigned_sources.pop(0)

                # 3.10) Heurística gulosa para agrupar sources e unir conjuntos de destinos
                while len(unsigned_sources) > 0:

                    next_points = [i for i in close_points[last_source] if i in unsigned_sources]
                    if len(next_points) > 0:
                        i = next_points[0]
                    else:
                        i = unsigned_sources[0]

                    # ATENÇÃO: Linha abaixo usa união de sets e depois len(...). A ideia é estimar cardinalidade de
                    # destinos distintos após adicionar o próximo i, multiplicada pelo #sources do pacote.
                    # Mantido conforme original.
                    if (len(filial_calls[-1]['sources']) + 1)*len(set(filial_calls[-1]['destinations']) | set(close_points[i])) > api_limit:
                        filial_calls.append({
                            'sources':[i],
                            'destinations':close_points[i]
                        })
                    else:
                        filial_calls[-1]['sources'].append(i)
                        filial_calls[-1]['destinations'] = list(set(filial_calls[-1]['destinations'] + close_points[i]))
                    
                    unsigned_sources.remove(i)

                # Indexar chamadas por nome legível
                calls[filial] = {
                    f'{filial} - {i + 1}':call
                    for i, call in enumerate(filial_calls)
                }

            # 3.11) Persistir plano de chamadas em JSON
            with open(model_data_folder + "chamadas_api.json", "w") as json_file:
                json.dump(calls, json_file, indent=4)
            
            # DataFrame acumulador dos resultados brutos; lista de chamadas já processadas
            result_df = pd.DataFrame()
            removed_calls = []

            print('Iniciando chamadas de API')

        else:
            # 3') Caminho de retomada: recarrega resultados brutos e marca chamadas já feitas
            print('Retomando chamadas de API')

            result_df = pd.read_parquet(model_data_folder + 'distance_matrix_raw.parquet')
            removed_calls = result_df['CALL'].unique().tolist()

        # 4) Carregar plano de chamadas e de-para
        with open(model_data_folder + "chamadas_api.json", "r") as json_file:
            calls = json.load(json_file)

        depara_point_id = pd.read_parquet(model_data_folder + 'depara_point_id.parquet')
        points_df = depara_point_id[['FILIAL','POINT_ID', 'LAT', 'LON']].drop_duplicates(subset=['POINT_ID','FILIAL']).copy()

        # Dicionário (FILIAL, POINT_ID) -> {'LAT':..., 'LON':...}
        coords_dict = points_df.groupby(['FILIAL', 'POINT_ID'])[['LAT', 'LON']].first().to_dict(orient='index')

        # Controle simples para espaçar chamadas (rate limiting)
        start = time.time()

        # 5) Loop principal de chamadas por filial/call
        for filial, filial_calls in calls.items():
            print(f'Iniciando chamadas para filial {filial}, modal {modal}')
            for call_name, call in tqdm(filial_calls.items()):
                if not(call_name in removed_calls):

                    # 5.1) Agrupar todas as coordenadas únicas desta call
                    locations_ids = list(set(call['sources'] + call['destinations']))

                    # Índices (no vetor 'locations') das sources e destinations
                    source_index = [i for i,point in enumerate(locations_ids) if point in call['sources']]
                    destination_index = [i for i,point in enumerate(locations_ids) if point in call['destinations']]

                    # Montar (lon, lat) para ORS
                    locations = [(coords_dict[filial, point]['LON'], coords_dict[filial, point]['LAT']) for point in locations_ids]

                    # 5.2) Rate limit básico (espera mínima entre chamadas)
                    end = time.time()
                    if end - start <= 1.6:
                        time.sleep(1.6 - (end - start))

                    start = time.time()

                    # 5.3) Requisitar matriz de distâncias e tempos à ORS
                    # Inicializando o cliente OpenRouteService
                    client = openrouteservice.Client(key=API_KEY)

                    try:
                        matrix = client.distance_matrix(
                            locations=locations,
                            sources=source_index,
                            destinations=destination_index,
                            profile=profile_dict[modal],
                            metrics=['distance', 'duration']  # Solicitar distância e duração
                        )
                    except:
                        # Retry simples em caso de erro transitório
                        time.sleep(3)
                        matrix = client.distance_matrix(
                            locations=locations,
                            sources=source_index,
                            destinations=destination_index,
                            profile=profile_dict[modal],
                            metrics=['distance', 'duration']  # Solicitar distância e duração
                        )

                    # 5.4) Converter resposta para DataFrame “long”
                    matrix_df = {'FILIAL':[], 'CALL':[],'POINT_ID_I':[], 'POINT_ID_J':[], 'DISTANCE':[], 'DURATION':[]}

                    for x, id_s in enumerate(source_index):
                        for y, id_d in enumerate(destination_index):
                            matrix_df['FILIAL'].append(filial)
                            matrix_df['CALL'].append(call_name)
                            matrix_df['POINT_ID_I'].append(locations_ids[id_s])
                            matrix_df['POINT_ID_J'].append(locations_ids[id_d])
                            matrix_df['DISTANCE'].append(matrix['distances'][x][y])
                            matrix_df['DURATION'].append(matrix['durations'][x][y])

                    matrix_df = pd.DataFrame(matrix_df)

                    # 5.5) Acumular e salvar resultado bruto incrementalmente por modal
                    result_df = pd.concat([result_df, matrix_df])
                    result_df.to_parquet(model_data_folder + f'{modal}_distance_matrix_raw.parquet', index=False)

        print('Tratando dados Obtidos')
        
        # 6) Tratamento de nulos e estimativas para pares não retornados pela API
        # ----------------------------------------------------------------------------
        # Passos:
        #   a) Criar matriz completa (todas combinações I–J por filial)
        #   b) Calcular DIST aproximada (km) via projeção planar com ajuste por latitude
        #   c) Obter razões médias DISTANCE/DIST e DURATION/DIST por origem e por filial
        #   d) Preencher nulos primeiro com média por origem; depois com média por filial

        # 6.a) Matriz completa com merge dos resultados (podem ter nulos)
        distance_matrix_df = (
            points_df
            .merge(points_df, on=['FILIAL'], how='left', suffixes=('_I', '_J'))
            .merge(result_df.drop(columns='CALL'), on=['FILIAL', 'POINT_ID_I', 'POINT_ID_J'], how='left')
        )

        # 6.b) DIST (km) aproximada para estimar densidade/velocidade média
        distance_matrix_df["DIST"] = np.sqrt(
            ((distance_matrix_df["LAT_J"] - distance_matrix_df["LAT_I"]) * 111.32) ** 2 +
            ((distance_matrix_df["LON_J"] - distance_matrix_df["LON_I"]) * 111.32 *
            np.cos(np.radians((distance_matrix_df["LAT_I"] + distance_matrix_df["LAT_J"]) / 2))) ** 2
        )

        # Razões auxiliares (quando há valores da API)
        distance_matrix_df['DISTANCE_PER_DIST'] = distance_matrix_df['DISTANCE']/distance_matrix_df['DIST']
        distance_matrix_df['DURATION_PER_DIST'] = distance_matrix_df['DURATION']/distance_matrix_df['DIST']

        # 6.c) Médias por origem (POINT_ID_I, FILIAL) e por FILIAL
        aux_df_1 = distance_matrix_df.groupby(['POINT_ID_I','FILIAL'])[['DISTANCE_PER_DIST', 'DURATION_PER_DIST']].mean().reset_index()
        aux_df_2 = distance_matrix_df.groupby(['FILIAL'])[['DISTANCE_PER_DIST', 'DURATION_PER_DIST']].mean().reset_index()

        # 6.d) Preenchimento em 2 estágios: por origem -> por filial
        distance_matrix_df = distance_matrix_df.drop(columns=['DISTANCE_PER_DIST', 'DURATION_PER_DIST']).merge(aux_df_1, on=['POINT_ID_I','FILIAL'], how='left')

        index_na = distance_matrix_df['DURATION'].isna()
        distance_matrix_df.loc[index_na, 'DURATION'] = distance_matrix_df.loc[index_na, 'DURATION_PER_DIST'] * distance_matrix_df.loc[index_na, 'DIST']
        distance_matrix_df.loc[index_na, 'DISTANCE'] = distance_matrix_df.loc[index_na, 'DISTANCE_PER_DIST'] * distance_matrix_df.loc[index_na, 'DIST']

        distance_matrix_df = distance_matrix_df.drop(columns=['DISTANCE_PER_DIST', 'DURATION_PER_DIST']).merge(aux_df_2, on=['FILIAL'], how='left')

        index_na = distance_matrix_df['DURATION'].isna()
        distance_matrix_df.loc[index_na, 'DURATION'] = distance_matrix_df.loc[index_na, 'DURATION_PER_DIST'] * distance_matrix_df.loc[index_na, 'DIST']
        distance_matrix_df.loc[index_na, 'DISTANCE'] = distance_matrix_df.loc[index_na, 'DISTANCE_PER_DIST'] * distance_matrix_df.loc[index_na, 'DIST']

        # Selecionar colunas finais (uma linha por par I–J por filial)
        distance_matrix_df = distance_matrix_df[['FILIAL', 'POINT_ID_I', 'POINT_ID_J', 'DISTANCE', 'DURATION']]

        print('Salvando Dados')
        # 7) Persistência final por modal (carro/a_pe)
        distance_matrix_df.to_parquet(model_data_folder + f'{modal}_distance_matrix.parquet', index=False)

        # Finalização amigável (modo não developer)
        if not developer:
            print('Processo finalizado, aperte enter para continuar')
            input()
