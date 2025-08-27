import pandas as pd
import numpy as np
import os
import xlwings as xw
import time
import folium
from ortools.constraint_solver import pywrapcp
from ortools.constraint_solver import routing_enums_pb2
from tqdm import tqdm
import sys
from collections import defaultdict
import win32com.client

pd.set_option('display.max_columns', None)


def close_excel_file_if_open(filename):
    """
    Fecha um arquivo do Excel se estiver aberto em alguma instância do Excel controlada via xlwings.
    Evita problemas de escrita/concorrência ao atualizar planilhas.
    """
    filename = os.path.basename(filename)
    try:
        for app in xw.apps:
            for wb in app.books:
                if wb.name == filename:
                    print(f"Fechando {filename}...")
                    wb.close()
                    time.sleep(1)
                    return True
    except Exception as e:
        print(f"Erro ao tentar fechar: {e}")
    return False



class Lote():
    """
    Modelo de roteirização (um lote/rota) com janelas de tempo.
    Usa OR-Tools (RoutingModel) para resolver um TSP/VRP com:
      - matriz de distância como custo de arco
      - dimensão de tempo (Time) para respeitar janelas
      - dimensão de distância (Distance) com limite máximo
    """
    def __init__(self):
        
        # -----------------------------
        # Conjuntos (nós/vertices)
        # -----------------------------
        self.clients = []   # lista de nós (clientes) a serem visitados
        self.vertices = []  # lista de nós incluindo base (definida em solve)

        # -----------------------------
        # Parâmetros globais do lote
        # -----------------------------
        self.route_cost = 10000   # custo fixo ao sair da base (penaliza abrir rota)
        self.base = 'BASE'        # rótulo do nó base
        self.max_time = 7         # tempo máximo (segundos) permitido para a rota (usado na dimensão Time)
        self.max_dist = 1000000   # distância máxima na dimensão Distance

        # -----------------------------
        # Parâmetros de nós
        # -----------------------------
        self.demand = {}       # demanda por nó (não utilizada neste solve 1-veículo)
        self.tw_start = {}     # início da janela de tempo por nó (em segundos)
        self.tw_end = {}       # fim da janela de tempo por nó (em segundos)
        self.service_time = {} # tempo de serviço por nó (em segundos)

        # -----------------------------
        # Parâmetros de arcos
        # -----------------------------
        self.distance = {}  # distância entre nós (m)
        self.time = {}      # tempo de deslocamento entre nós (s)

        self.infeasible_clients = []    # lista de clientes inviáveis (tempo ida+serviço+volta > max_time)  
    

    def solve(self, time_limit=600, verbose=False, base='BASE'):
        """
        Resolve a roteirização para um único veículo.
        time_limit: limite de processamento (s)
        verbose:   ativa logs de busca do OR-Tools
        base:      nome do nó base (pode ser 'BASE' ou outro; altera custo da matriz de distância)
        """

        if verbose:
            print('Iniciando Parametrização')

        # =========================
        # 0) Parametrização
        # =========================
        # 0.1) Conjuntos de nós
        clients = self.clients
        nodes = [self.base] + clients  # base é o índice 0 na matriz

        # 0.2) Distâncias e Tempos (filtra apenas nós do problema)
        dist = {i:{j:int(v) for j,v in v_i.items() if j in nodes} for i,v_i in self.distance.items() if i in nodes}
        t = {i:{j:int(v) for j,v in v_i.items() if j in nodes} for i,v_i in self.time.items() if i in nodes}

        # 0.3) Janelas de tempo relativas ao início da base
        e = {i: int(max(v - self.tw_start[self.base], 0))  for i,v in self.tw_start.items() if i in nodes}
        l = {i: int((v - self.tw_start[self.base])) for i,v in self.tw_end.items() if i in nodes}

        # 0.5) Tempos de serviço (0 na base)
        s = {i:int(v) for i,v in self.service_time.items() if i in clients} | {self.base:0}

        # =========================
        # 1) Estruturas de dados no formato do OR-Tools
        # =========================
        data = {}

        # 1.1) Matriz de distâncias (aplica penalidade de saída da base quando base == 'BASE')
        if base == 'BASE':
            data['distance_matrix'] = [
                [
                    dist[i][j] + (i == self.base)*self.route_cost
                    for j in nodes
                ]
                for i in nodes
            ]
        else:
            data['distance_matrix'] = [
                [
                    dist[i][j]
                    for j in nodes
                ]
                for i in nodes
            ]

        # 1.2) Matriz de tempos (tempo de deslocamento + tempo de serviço do nó de origem)
        data['time_matrix'] = [
            [
                t[i][j] + s[i]
                for j in nodes
            ]
            for i in nodes
        ]

        # 1.3) Janelas de tempo absolutas (relativas à base)
        data['time_windows'] = [(e[i], l[i]) for i in nodes]

        # 1.4) Número de veículos e depósito (aqui sempre 1 veículo)
        data['num_vehicles'] = 1
        data['depot'] = 0

        # =========================
        # 2) Construção do modelo OR-Tools
        # =========================
        # 2.1) Gerenciador de índices
        manager = pywrapcp.RoutingIndexManager(len(data['distance_matrix']),
                                            data['num_vehicles'], data['depot'])

        # 2.2) Modelo de roteamento
        routing = pywrapcp.RoutingModel(manager)
        
        # 2.3) Callback de custo de arco (distância)
        def distance_callback(from_index, to_index):
            from_node = manager.IndexToNode(from_index)
            to_node = manager.IndexToNode(to_index)
            return data['distance_matrix'][from_node][to_node]

        transit_distance_callback_index = routing.RegisterTransitCallback(distance_callback)
        routing.SetArcCostEvaluatorOfAllVehicles(transit_distance_callback_index)

        # 2.4) Callback de tempo (separado da distância)
        def time_callback(from_index, to_index):
            from_node = manager.IndexToNode(from_index)
            to_node = manager.IndexToNode(to_index)
            return data['time_matrix'][from_node][to_node]

        transit_time_callback_index = routing.RegisterTransitCallback(time_callback)

        # 2.5) Dimensão de Tempo (com janelas)
        routing.AddDimension(
            transit_time_callback_index,
            self.max_time,  # tempo de espera máximo permitido
            self.max_time,  # tempo máximo de percurso
            False,          # False: permite “espera” (Time Slack) dentro da janela
            'Time')

        time_dimension = routing.GetDimensionOrDie('Time')

        # 2.6) Dimensão de Distância (para limitar percurso total)
        routing.AddDimension(
            transit_distance_callback_index,
            0,              # sem folga
            self.max_dist,  # distância máxima por rota
            True,           # rota deve começar/terminar no depósito
            'Distance'
        )

        distance_dimension = routing.GetDimensionOrDie('Distance')

        # 2.7) Limite máximo de distância por veículo
        for vehicle_id in range(data['num_vehicles']):
            distance_dimension.CumulVar(routing.End(vehicle_id)).SetMax(self.max_dist)

        # 2.8) Janelas de tempo por nó (exceto base, feita abaixo)
        for location_idx, time_window in enumerate(data['time_windows']):
            if location_idx == 0:
                continue
            index = manager.NodeToIndex(location_idx)
            time_dimension.CumulVar(index).SetRange(int(time_window[0]), int(time_window[1]))

        # 2.9) Janela de tempo do depósito
        depot_idx = manager.NodeToIndex(data['depot'])
        time_dimension.CumulVar(depot_idx).SetRange(e[self.base], l[self.base])

        # 2.10) Estratégia/limite de busca
        search_parameters = pywrapcp.DefaultRoutingSearchParameters()
        search_parameters.first_solution_strategy = (
            routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC)
        search_parameters.time_limit.FromSeconds(time_limit)
        search_parameters.log_search = verbose

        # =========================
        # 3) Solução
        # =========================
        if verbose:
            print('Resolvendo Problema')

        solution = routing.SolveWithParameters(search_parameters)

        # =========================
        # 4) Extração da solução
        # =========================
        solution_dict = {}
        for vehicle_id in range(data['num_vehicles']):
            route_dict = {}
            index = routing.Start(vehicle_id)
            arc_count = 0

            if routing.IsEnd(solution.Value(routing.NextVar(index))):
                continue  # veículo não utilizado

            while not routing.IsEnd(index):
                from_node = manager.IndexToNode(index)
                next_index = solution.Value(routing.NextVar(index))
                to_node = manager.IndexToNode(next_index)
                i = nodes[from_node]
                j = nodes[to_node]

                arc_data = {
                    'arc': (i, j),
                    'dist': self.distance[i][j],
                    'time': self.time[i][j],
                    'service_time': self.service_time.get(i, 0),  # 0 se for depósito
                    'demand': self.demand.get(i, 0)               # 0 se for depósito
                }
                route_dict[arc_count] = arc_data

                arc_count += 1
                index = next_index

            solution_dict[vehicle_id] = route_dict

        return solution_dict

    
    def remove_infesible_points(self):
        """
        Filtra clientes inviáveis: ida + serviço + volta excedem self.max_time.
        Atualiza self.infeasible_clients e reduz self.clients.
        """
        self.infeasible_clients = [client for client in self.clients 
                        if self.time[self.base][client] + self.time[client][self.base] + self.service_time[client] > self.max_time]
        
        self.clients = [client for client in self.clients 
                        if self.time[self.base][client] + self.time[client][self.base] + self.service_time[client] <= self.max_time]

def roteirizar(lote_df, livro, filial, time_limit, distance_matrix, periodo, base_abastecedor='BASE', base_point_id=''):
    """
    Monta e resolve um Lote (1 veículo) para um conjunto de parceiros (lote_df),
    retornando um DataFrame com as visitas resultantes para o 'livro' especificado.
    - base_abastecedor: 'BASE' (fantasma) ou nome de um parceiro que atua como base real
    - base_point_id: POINT_ID da base quando base_abastecedor != 'BASE'
    """

    # -----------------------------
    # Constrói dicionários de distância/tempo a partir da matriz da filial
    # -----------------------------
    dist_dict = distance_matrix[distance_matrix['FILIAL'] == filial].set_index(['POINT_ID_I','POINT_ID_J'])['DISTANCE'].to_dict()
    time_dict = distance_matrix[distance_matrix['FILIAL'] == filial].set_index(['POINT_ID_I','POINT_ID_J'])['DURATION'].to_dict()

    # -----------------------------
    # Parâmetros por nó (serviço/entrada) e janelas
    # -----------------------------
    tempo_servico = lote_df.set_index('PARCEIRO')['TEMPO_SERVICO'].to_dict() | {base_abastecedor: 0}
    tempo_entrada = lote_df.set_index('PARCEIRO')['TEMPO_DE_ENTRADA'].to_dict() | {base_abastecedor: 0}
    point = lote_df.set_index('PARCEIRO')['POINT_ID'].to_dict() | {base_abastecedor: base_point_id}
    inicio = lote_df.set_index('PARCEIRO')['INICIO'].to_dict() | {base_abastecedor: 0}
    fim = lote_df.set_index('PARCEIRO')['FIM'].to_dict() | {base_abastecedor: int(198*3600)}

    # -----------------------------
    # Instancia o Lote e preenche parâmetros
    # -----------------------------
    lote = Lote()

    # Conjuntos (nós clientes e base)
    lote.clients = lote_df['PARCEIRO'].to_list()
    lote.vertices = [base_abastecedor] + lote.clients

    # Parâmetros globais do lote (tempo/distâncias grandes para não limitar reroteirização)
    lote.route_cost = 1000000
    lote.base = base_abastecedor
    lote.max_time = int(198*3600)          # 198h ~ janela ampliada
    lote.max_dist = 10000000 + lote.route_cost

    # Janelas/serviço
    lote.tw_start = inicio
    lote.tw_end = fim
    lote.service_time = tempo_servico

    # Matrizes de distância/tempo:
    # - Se base for 'BASE', mantemos custo 0 aos arcos que envolvem a base fantasma
    # - Se base real, usamos dist/tempo completos entre todos os vértices
    if base_abastecedor == 'BASE':
        lote.distance = {i:{j:dist_dict[point[i], point[j]] for j in lote.clients} | {base_abastecedor:0} for i in lote.clients} | {base_abastecedor:{j:0 for j in lote.vertices}}
        lote.time = {i:{j:time_dict[point[i], point[j]] for j in lote.clients} | {base_abastecedor:0} for i in lote.clients} | {base_abastecedor:{j:0 for j in lote.vertices}}
    else:
        lote.distance = {i:{j:dist_dict[point[i], point[j]] for j in lote.vertices} for i in lote.vertices}
        lote.time = {i:{j:time_dict[point[i], point[j]] for j in lote.vertices} for i in lote.vertices}

    # Adiciona tempo de entrada ao tempo de deslocamento de destino (troca de parceiro)
    for i in lote.vertices:
        for j in lote.vertices:
            lote.time[i][j] = lote.time[i][j] + tempo_entrada[j]

    # Resolve o lote (OR-Tools) e estrutura o DataFrame de visitas
    result_dict = lote.solve(time_limit=time_limit, verbose=False, base=base_abastecedor)

    livros_df = {'FILIAL':[], 'PERIODO':[], 'LIVRO':[], 'VISITA':[], 'PARCEIRO':[], 'DIST':[], 'TEMPO_DESLOCAMENTO':[], 'TEMPO_DE_SERVICO':[]}

    for visitas in result_dict.values():
        livro_name = livro
        for visita, arco in visitas.items():
            livros_df['FILIAL'].append(filial)
            livros_df['PERIODO'].append(periodo)
            livros_df['LIVRO'].append(livro_name)
            livros_df['VISITA'].append(visita + 1)
            livros_df['PARCEIRO'].append(arco['arc'][1])
            livros_df['DIST'].append(arco['dist'])
            livros_df['TEMPO_DESLOCAMENTO'].append(arco['time'])
            livros_df['TEMPO_DE_SERVICO'].append(tempo_servico[arco['arc'][1]])
        
    livros_df = pd.DataFrame(livros_df)

    return livros_df


def std_codes(code):
    """
    Padroniza códigos/string vindos do Excel, removendo quebras e zeros à esquerda em numéricos.
    """
    if str(code).replace('.','').isdigit():
        return(str(int(code))).replace('_x000D_\n', '').replace('\n', '')
    else:
        return str(code).replace('_x000D_\n', '').replace('\n', '')
    
    
def main(atualizar_rotas = False):
    """
    Recalcula (“reroteiriza”) rotas após ajustes manuais:
      1) Lê “Livros” já existentes (Excel)
      2) (Opcional) Recalcula a ordem de parceiros de cada livro usando OR-Tools
      3) Recalcula distâncias/tempos por modal (carro / a pé)
      4) Reaplica tempo de entrada na troca de parceiro
      5) Exporta um Excel consolidado com resultados ajustados
    """
    # =========================
    # 1) Pastas e arquivos
    # =========================
    current_script_directory = os.path.dirname(os.path.abspath(__file__))
    main_dir = '/'.join(current_script_directory.split('\\')[:-1]) + '/'
    model_data_folder = main_dir + 'Dados Intermediários/'
    input_folder = main_dir + 'Dados de Input/'
    adjusted_livros_file = 'Piloto Reroteirizado.xlsx'
    
    # Descobre workbook (.xlsm) principal do projeto
    for file in os.listdir(main_dir):
        if len(file) > 4:
            if file[-5:] == '.xlsm':
                workbook = file
                
    close_excel_file_if_open(main_dir + workbook)
    # wb não é aberto (somente leitura do xlsm na linha adiante)
  

    # =========================
    # 2) Leitura de bases
    # =========================
    # 2.1) De-para de pontos (POINT_ID e coordenadas)
    depara_point_id = pd.read_parquet(model_data_folder + 'depara_point_id.parquet')[['FILIAL', 'PARCEIRO', 'POINT_ID', 'LAT', 'LON']]
    depara_point_id['PARCEIRO'] = depara_point_id['PARCEIRO'].apply(std_codes)

    # 2.2) Matrizes de distância (carro e a pé)
    distance_matrix = pd.read_parquet(model_data_folder + 'carro_distance_matrix.parquet')
    distance_matrix['DISTANCE'] = distance_matrix['DISTANCE'].round(0).astype(int)
    distance_matrix['DURATION'] = (distance_matrix['DURATION']*1.05).round(0).astype(int)  # fator de trânsito

    a_pe_distance_matrix = pd.read_parquet(model_data_folder + 'a_pe_distance_matrix.parquet')
    a_pe_distance_matrix['DISTANCE'] = a_pe_distance_matrix['DISTANCE'].round(0).astype(int)
    a_pe_distance_matrix['DURATION'] = (a_pe_distance_matrix['DURATION']*1.05).round(0).astype(int)

    # 2.3) Parceiros (janelas de tempo e tempo de entrada)
    parceiros_df = pd.read_excel(model_data_folder + 'Dados.xlsx', sheet_name='parceiros')
    parceiros_df['PARCEIRO'] = parceiros_df['PARCEIRO'].apply(std_codes)

    # 2.4) Livros originais (aba “Livros” do workbook principal)
    livros_df = pd.read_excel(main_dir + workbook, skiprows=5, usecols='B:M', sheet_name='Livros')
    
    
    # =========================
    # 3) Preparação de janelas de tempo
    # =========================
    parceiros_df_aux = parceiros_df.copy()
    index = parceiros_df_aux['INICIO_FUNCIONAMENTO'].notna()
    parceiros_df_aux.loc[index, 'INICIO'] = parceiros_df_aux.loc[index, 'INICIO_FUNCIONAMENTO'].astype(str).apply(lambda x: int(x.split(':')[0])*3600 + int(x.split(':')[1])*60).astype(int)
    index = parceiros_df_aux['FIM_FUNCIONAMENTO'].notna()
    parceiros_df_aux.loc[index, 'FIM'] = (parceiros_df_aux.loc[index, 'FIM_FUNCIONAMENTO']
                                        .astype(str).apply(lambda x: int(x.split(':')[0])*3600 + 
                                                            int(x.split(':')[1])*60).astype(int))

    # Ajuste para janelas que passam da meia-noite (fim < início)
    index = parceiros_df_aux['INICIO'] > parceiros_df_aux['FIM']
    parceiros_df_aux.loc[index, 'FIM'] = parceiros_df_aux.loc[index, 'FIM'] + 48*3600

    # Normaliza referência temporal (offset pelo menor início)
    min_inicio = parceiros_df_aux['INICIO'].min()
    parceiros_df_aux['INICIO'] = parceiros_df_aux['INICIO'].fillna(min_inicio) - min_inicio
    parceiros_df_aux['FIM'] = parceiros_df_aux['FIM'] - min_inicio

    # Completa FIM ausente com 48h
    max_fim = parceiros_df_aux['FIM'].max()
    index = parceiros_df_aux['FIM'].isna()
    parceiros_df_aux.loc[index, 'FIM'] = np.maximum(max_fim, 48*3600).astype(int)

    # Mantém apenas campos necessários + converte tempos para segundos
    parceiros_df_aux['INICIO'] = parceiros_df_aux['INICIO'].astype(int)
    parceiros_df_aux['FIM'] = parceiros_df_aux['FIM'].astype(int)
    parceiros_df_aux = parceiros_df_aux[['FILIAL', 'PARCEIRO', 'TEMPO_DE_ENTRADA_MIN', 'INICIO', 'FIM']].rename(columns={'TEMPO_DE_ENTRADA_MIN':'TEMPO_DE_ENTRADA'})
    parceiros_df_aux['TEMPO_DE_ENTRADA'] = (parceiros_df_aux['TEMPO_DE_ENTRADA']*60).astype(int)

    # =========================
    # 4) (Opcional) Reroteirização completa com OR-Tools
    # =========================
    if atualizar_rotas:
        # De/para de POINT_ID para os parceiros do relatório de livros
        livros_df_aux2 = (
            livros_df
            .merge(depara_point_id[['PARCEIRO', 'FILIAL', 'POINT_ID']],
                left_on=['Parceiro', 'Filial'],
                right_on=['PARCEIRO', 'FILIAL'],
                how='left')
            .drop(columns=['PARCEIRO', 'FILIAL']))

        result_df = pd.DataFrame()

        # Para cada livro, recalcula a ordem ótima dos parceiros via OR-Tools (um veículo)
        for livro in tqdm(livros_df_aux2['Livro'].unique(), file=sys.stdout):

            dfAux = livros_df_aux2[livros_df_aux2['Livro'] == livro].copy()
            # abastecedor_definido = dfAux['ABASTECEDOR DEFINIDO'].unique()[0]  # (não utilizado)
            modal = dfAux['Modal de Transporte'].unique()[0]

            # Consolida tempo de serviço por parceiro (no livro/dia/filial)
            dfAux = (
                dfAux
                .assign(TEMPO_SERVICO=livros_df_aux2['Tempo de Serviço (min)']*60)
                .assign(PARCEIRO=livros_df_aux2['Parceiro'])
                .groupby(['PARCEIRO', 'POINT_ID', 'Dia', 'Filial'])
                .agg({'TEMPO_SERVICO':'sum'})
                .reset_index()
                .merge(parceiros_df_aux,
                    on='PARCEIRO',
                    how='left')
                .copy()
            )
            filial = dfAux['Filial'].unique()[0]
            periodo = dfAux['Dia'].unique()[0]
            base_abastecedor = "BASE"   # base fantasma
            base_point_id = ''
            time_limit = 200

            # Escolhe matriz de deslocamento conforme modal do livro
            if modal == 'A pé':
                distance_matrix_aux = a_pe_distance_matrix.copy()
            else:
                distance_matrix_aux = distance_matrix.copy()

            # Reroteiriza e agrega ao resultado
            output = roteirizar(dfAux, livro, filial, time_limit, distance_matrix_aux, periodo, base_abastecedor, base_point_id)
            result_df = pd.concat([result_df, output])

        # Mapeia nova ordem de visitas por (Filial, Livro, Parceiro)
        result_df_aux = result_df[['FILIAL', 'LIVRO', 'PARCEIRO', 'VISITA']].copy()
        
    # =========================
    # 5) Recalcula métricas de deslocamento/serviço por visita (com/sem reroteirização)
    # =========================
    livros_df_aux = livros_df.copy()
    livros_df_aux['Patrimônio'] = livros_df_aux['Patrimônio'].astype(str)

    # 5.1) Se houve reroteirização, substitui a ordem de “# Visita”
    if atualizar_rotas:
        livros_df_aux = (
            livros_df_aux
            .merge(result_df_aux.
                rename(columns={'FILIAL':'Filial',
                                'LIVRO': 'Livro',
                                'PARCEIRO': 'Parceiro'}),
                                on=['Filial', 'Livro', 'Parceiro'],
                                how='left')
            .sort_values(by=['Filial', 'Dia', 'Livro', 'VISITA'])
            .reset_index(drop=True)
            .reset_index()
            )

        # Reatribui # Visita por livro (1..N)
        livros_df_aux['# Visita'] = livros_df_aux.groupby(['Livro'])['index'].transform(lambda x: x.rank(method='dense').astype(int))
        livros_df_aux.drop(columns=['index', 'VISITA'], inplace=True)

    # Ordena visitas e cria índice auxiliar (para pares consecutivos)
    livros_df_aux = livros_df_aux.sort_values(by=['Filial', 'Dia', 'Livro', '# Visita']).reset_index(drop=True).reset_index()

    # Junta POINT_ID de destino do parceiro da linha e calcula origem como o anterior
    livros_df_aux = livros_df_aux.merge(depara_point_id
                                    .rename(columns={'PARCEIRO':'Parceiro',
                                                    'FILIAL':'Filial',
                                                    'POINT_ID':'POINT_ID_J'}),
                                    on=['Parceiro', 'Filial'],
                                    how='left')
    livros_df_aux['Patrimônio'] = livros_df_aux['Patrimônio'].astype(str)

    # Cria POINT_ID_I (origem = destino anterior dentro do mesmo livro)
    livros_df_aux['POINT_ID_I'] = livros_df_aux['POINT_ID_J'].shift(1)
    livros_df_aux.loc[livros_df_aux['# Visita'] == 1, 'POINT_ID_I'] = None

    # 5.2) Junta distâncias/tempos por modal (carro e a pé) para cada arco (I->J)
    livros_df_aux = (livros_df_aux
                    .merge(distance_matrix
                        .rename(columns={'FILIAL':'Filial',
                                            'DISTANCE': 'CAR_DISTANCE',
                                            'DURATION': 'CAR_DURATION'}),
                        on=['Filial', 'POINT_ID_I', 'POINT_ID_J'],
                        how='left')
                    .assign(PUB_TRANS_DURATION=lambda x: x['CAR_DURATION']+600)  # proxy de transporte público (não utilizado adiante)
                    .merge(a_pe_distance_matrix
                    .rename(columns={'FILIAL':'Filial',
                                    'DISTANCE': 'WALK_DISTANCE',
                                    'DURATION': 'WALK_DURATION'}),
                                    on=['Filial', 'POINT_ID_I', 'POINT_ID_J'],
                                    how='left'))

    # Seleciona distância/tempo conforme modal do livro
    index_ape = (livros_df_aux['Modal de Transporte'] == 'A pé')
    livros_df_aux['DISTANCE'] = livros_df_aux['CAR_DISTANCE']
    livros_df_aux['DURATION'] = livros_df_aux['CAR_DURATION']

    livros_df_aux.loc[index_ape, 'DISTANCE'] = livros_df_aux['WALK_DISTANCE']
    livros_df_aux.loc[index_ape, 'DURATION'] = livros_df_aux['WALK_DURATION']

    # Garante 0 em lacunas e remove colunas auxiliares de cálculo
    livros_df_aux = (
        livros_df_aux
        .fillna({'DISTANCE': 0, 'DURATION':0})
        .drop(columns=['CAR_DISTANCE', 'CAR_DURATION', 'WALK_DISTANCE', 'WALK_DURATION']))

    # 5.3) Aplica tempo de entrada quando troca de parceiro (primeira visita do livro também conta)
    livros_df_aux['TEMPO_DE_ENTRADA'] = 600  # 10 min padrão
    livros_df_aux['PARCEIRO_I'] = livros_df_aux['Parceiro'].shift(1)
    livros_df_aux['REMOVER_TEMPO_DE_ENTRADA'] = ((livros_df_aux['PARCEIRO_I'] == livros_df_aux['Parceiro']) &
                                                (livros_df_aux['# Visita'] != 1))
    livros_df_aux.loc[livros_df_aux['REMOVER_TEMPO_DE_ENTRADA'], 'TEMPO_DE_ENTRADA'] = 0
    livros_df_aux['DURATION'] = livros_df_aux['DURATION'] + livros_df_aux['TEMPO_DE_ENTRADA']

    # 5.4) Conversões finais para relatório (km / min)
    livros_df_aux['Distância (km)'] = livros_df_aux['DISTANCE']/1000
    livros_df_aux['Tempo de Deslocamento (min)'] = livros_df_aux['DURATION']/60

    # Remove colunas técnicas e escreve Excel final
    livros_df_aux.drop(columns=['POINT_ID_J', 'LAT', 'LON', 'POINT_ID_I', 'PUB_TRANS_DURATION', 'DISTANCE', 'DURATION',
                                'TEMPO_DE_ENTRADA', 'PARCEIRO_I', 'REMOVER_TEMPO_DE_ENTRADA', 'index'],
                        inplace=True)

    # =========================
    # 6) Exporta resultado para Excel (planilha única)
    # =========================
    livros_df_aux.to_excel(main_dir + adjusted_livros_file, index=False)
    livros_df_aux
