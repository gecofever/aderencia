"""Dashboard PCP - Indicador On Time (acabamento).

Versão com P0 aplicado:
 1. Datas de lançamento e de limite são normalizadas (sem hora) antes da comparação.
 2. OP sem Data Stábil recebe status próprio ('Sem Data') em vez de 'Atrasado'.
 3. Calendário: status_6x1 convertido com to_numeric e o conjunto de dias sem
    produção calculado uma única vez (blocos contíguos), não por linha.
 4. Cálculo de prazo vetorizado, com aviso quando o prazo sai da cobertura do
    calendário.
"""

import pandas as pd
import numpy as np
import streamlit as st

# ==============================================================================
# 1. CONFIGURAÇÕES
# ==============================================================================

st.set_page_config(page_title="Dashboard PCP - On Time", layout="wide")

pd.set_option('display.float_format', lambda x: '%.2f' % x)

PATH_SAIDA_PRENSAS = r"M:\CGE-Pcp\Astro\Análises Programação\1- Prensas\Controle Horários de prensa\Saída de prensas.xlsx"
PATH_YMB51 = r"M:\CGE-Pcp\Pilar Planejamento\06 - Indicadores\3. Base de Dados\Entrada de CD\10.Outubro.XLSX"
PATH_CALENDARIO = r"M:\CGE-Pcp\Astro\Análises Programação\1- Prensas\Controle Horários de prensa\Saída de prensas.xlsx"

COLUNAS_ALVO = ['OP Acabado', 'Data Stábil', 'Qtd. Par Real', 'Cód. Artigo', 'Desc. Artigo']
COL_YMB51 = ['Ordem', 'Material', 'Qtd.  UM registro', 'Data de lançamento']
COL_CALENDARIO = ['Data', '6x1']

PRAZO_DIAS = 3          # dias corridos após a Data Stábil
STATUS_ON_TIME = 'On Time'
STATUS_ATRASADO = 'Atrasado'
STATUS_PENDENTE = 'Pendente'
STATUS_SEM_DATA = 'Sem Data'

# ==============================================================================
# 2. FUNÇÕES DE PROCESSAMENTO
# ==============================================================================

def tratar_nomes(df):
    if df is None:
        return None
    renames_prensas = {'OP Acabado': 'op_acabado', 'Data Stábil': 'data_stabil', 'Qtd. Par Real': 'qtd_par_real',
                       'Cód. Artigo': 'cod_artigo', 'Desc. Artigo': 'desc_artigo'}
    renames_acab1 = {'Ordem': 'op_acabado', 'Material': 'cod_artigo', 'Qtd.  UM registro': 'qtd_par_real',
                     'Data de lançamento': 'data_lanca_cd'}
    if 'OP Acabado' in df.columns:
        return df.rename(columns=renames_prensas)
    if 'Ordem' in df.columns:
        return df.rename(columns=renames_acab1)
    return df


@st.cache_data(ttl=600)
def carregar_calendario_previsto(path: str = PATH_CALENDARIO):
    """Lê a aba de calendário e devolve data_prevista (normalizada) + status_6x1 numérico."""
    try:
        df_cal = pd.read_excel(path, sheet_name='Balanceado', usecols=COL_CALENDARIO,
                               engine='openpyxl', header=2)
        df_cal = df_cal.rename(columns={'Data': 'data_prevista', '6x1': 'status_6x1'})
        df_cal['data_prevista'] = pd.to_datetime(df_cal['data_prevista'], errors='coerce').dt.normalize()
        # P0.3: evita que '0' gravado como texto deixe de ser reconhecido como dia sem produção
        df_cal['status_6x1'] = pd.to_numeric(df_cal['status_6x1'], errors='coerce')
        df_cal = df_cal.dropna(subset=['data_prevista'])
        return df_cal
    except Exception as e:
        st.error(f"Erro ao ler a planilha de calendário ({path}): {e}")
        return None


def _blocos_sem_producao(df_cal):
    """Devolve (dias sem produção, início e fim de cada bloco contíguo), em dias desde a epoch."""
    vazio = np.array([], dtype='int64')
    if df_cal is None or df_cal.empty:
        return vazio, vazio, vazio

    sem_6x1 = df_cal.loc[df_cal['status_6x1'].fillna(0).eq(0), 'data_prevista']
    if sem_6x1.empty:
        return vazio, vazio, vazio

    dias = np.sort(sem_6x1.to_numpy().astype('datetime64[D]').astype('int64'))
    dias = np.unique(dias)
    novo_bloco = np.r_[True, np.diff(dias) > 1]
    pos_inicios = np.flatnonzero(novo_bloco)
    inicios = dias[pos_inicios]
    fins = np.r_[dias[pos_inicios[1:] - 1], dias[-1]]   # último dia de cada bloco
    return dias, inicios, fins


def calcular_prazo_operacional(data_stabil: pd.Series, df_cal) -> pd.Series:
    """Prazo = Data Stábil + PRAZO_DIAS, empurrado para o primeiro dia com produção.

    Vetorizado: uma única busca por linha (searchsorted), sem reconstruir o
    conjunto de dias sem produção para cada OP.
    """
    limite = (data_stabil + pd.Timedelta(days=PRAZO_DIAS)).dt.normalize()

    dias_sem_prod, inicios, fins = _blocos_sem_producao(df_cal)
    if dias_sem_prod.size == 0:
        return limite

    validos = limite.notna().to_numpy()
    if not validos.any():
        return limite

    dias = limite.to_numpy().astype('datetime64[D]').astype('int64')
    pos = np.clip(np.searchsorted(dias_sem_prod, dias), 0, dias_sem_prod.size - 1)
    cai_em_parada = validos & (dias_sem_prod[pos] == dias)

    if cai_em_parada.any():
        bloco = np.searchsorted(inicios, dias[cai_em_parada], side='right') - 1
        dias = dias.copy()
        dias[cai_em_parada] = fins[bloco] + 1          # primeiro dia após o bloco de parada

    resultado = pd.Series(pd.to_datetime(dias, unit='D'), index=data_stabil.index)
    resultado[~validos] = pd.NaT
    return resultado


def classificar_status(df: pd.DataFrame) -> pd.Series:
    """Classifica cada OP. 'Sem Data' não entra no denominador do indicador."""
    status = pd.Series(index=df.index, dtype='object')
    sem_data = df['data_stabil'].isna()
    sem_lancamento = df['data_lanca_cd'].isna()

    status[sem_data] = STATUS_SEM_DATA
    status[~sem_data & sem_lancamento] = STATUS_PENDENTE

    concluidas = (~sem_data) & (~sem_lancamento)
    if concluidas.any():
        entrou = df.loc[concluidas, 'data_lanca_cd'].dt.normalize()
        prazo = df.loc[concluidas, 'data_limite_ontime'].dt.normalize()
        status[concluidas] = np.where(entrou <= prazo, STATUS_ON_TIME, STATUS_ATRASADO)

    return status


@st.cache_data(ttl=600)
def processar_base_integrada(path_prensas: str = PATH_SAIDA_PRENSAS,
                            path_cd: str = PATH_YMB51,
                            path_calendario: str = PATH_CALENDARIO):
    try:
        df_p1 = pd.read_excel(path_prensas, sheet_name='Stábil UGB 1', usecols=COLUNAS_ALVO,
                              engine='openpyxl', header=2)
        df_p2 = pd.read_excel(path_prensas, sheet_name='Stábil UGB 2', usecols=COLUNAS_ALVO,
                              engine='openpyxl', header=2)
        df_acab1 = pd.read_excel(path_cd, sheet_name='Sheet1', usecols=COL_YMB51,
                                 engine='openpyxl', header=0)

        df_p1 = tratar_nomes(df_p1)
        df_p2 = tratar_nomes(df_p2)
        df_acab1 = tratar_nomes(df_acab1)

        # 1. CONSOLIDAÇÃO PRENSAS
        df_prensa_final = pd.concat([df_p1, df_p2], ignore_index=True).dropna(subset=['op_acabado'])
        df_prensa_final['op_acabado'] = df_prensa_final['op_acabado'].astype('Int64')
        df_prensa_final['cod_artigo'] = df_prensa_final['cod_artigo'].astype('Int64')
        df_prensa_final['data_stabil'] = pd.to_datetime(df_prensa_final['data_stabil'],
                                                        errors='coerce').dt.normalize()

        df_prensas_consolidado = df_prensa_final.groupby('op_acabado', as_index=False).agg({
            'data_stabil': 'min', 'qtd_par_real': 'sum', 'cod_artigo': 'first', 'desc_artigo': 'first'
        })

        # 2. CONSOLIDAÇÃO ACABAMENTO
        df_acab1 = df_acab1.dropna(subset=['op_acabado'])
        df_acab1['op_acabado'] = df_acab1['op_acabado'].astype('Int64')
        df_acab1['data_lanca_cd'] = pd.to_datetime(df_acab1['data_lanca_cd'], errors='coerce')

        df_acab_consolidado = df_acab1.groupby('op_acabado', as_index=False).agg({
            'data_lanca_cd': 'max', 'qtd_par_real': 'sum'
        })

        # 3. JOIN INTEGRADO
        df_final = pd.merge(df_prensas_consolidado, df_acab_consolidado, on='op_acabado',
                            how='left', suffixes=('_prensa', '_acabamento'))

        # 4. PRAZO E STATUS
        df_cal = carregar_calendario_previsto(path_calendario)
        if df_cal is not None and not df_cal.empty:
            df_final['data_limite_ontime'] = calcular_prazo_operacional(df_final['data_stabil'], df_cal)
            df_final['status_entrega'] = classificar_status(df_final)

            ultima_data_cal = df_cal['data_prevista'].max()
            fora_cobertura = int((df_final['data_limite_ontime'] > ultima_data_cal).sum())
            if fora_cobertura:
                st.warning(
                    f"{fora_cobertura} OP(s) com prazo depois de {ultima_data_cal:%d/%m/%Y}, "
                    "que é a última data do calendário. Atualize a planilha de calendário."
                )
        else:
            df_final['data_limite_ontime'] = pd.NaT
            df_final['status_entrega'] = 'Sem Calendário'

        return df_final
    except Exception as e:
        st.error(f"Erro ao processar as bases: {e}")
        return None


# ==============================================================================
# 3. INTERFACE
# ==============================================================================

st.title("📊 Painel de Controle de Produção - Indicador On Time")

df_base = processar_base_integrada()
df_calendario = carregar_calendario_previsto()

if df_base is not None:
    st.sidebar.header("🔍 Buscar por Artigo")
    busca_artigo = st.sidebar.number_input("Digite o Código do Artigo (ou 0 para todos):",
                                           min_value=0, value=0, step=1)

    if busca_artigo != 0:
        df_filtrado = df_base[df_base['cod_artigo'] == busca_artigo].copy()
        st.subheader(f"🎯 Resultados para o Artigo: {busca_artigo}")
    else:
        df_filtrado = df_base.copy()
        st.subheader("🌐 Visão Geral da Fábrica (Todos os Artigos)")

    contagem = df_filtrado['status_entrega'].value_counts()
    total_planilhadas = len(df_filtrado)
    total_on_time = int(contagem.get(STATUS_ON_TIME, 0))
    total_atrasado = int(contagem.get(STATUS_ATRASADO, 0))
    total_pendente = int(contagem.get(STATUS_PENDENTE, 0))
    total_sem_data = int(contagem.get(STATUS_SEM_DATA, 0))

    # Denominador: apenas OPs com prazo calculável e desfecho conhecido
    ops_concluidas = total_on_time + total_atrasado
    percentual_on_time = (total_on_time / ops_concluidas * 100) if ops_concluidas > 0 else None
    texto_percentual = f"{percentual_on_time:.1f}%" if percentual_on_time is not None else "n/d"

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total de OPs", f"{total_planilhadas}")
    c2.metric("On Time", f"{total_on_time}")
    c3.metric("Atrasado", f"{total_atrasado}", delta_color="inverse")
    c4.metric("Pendente", f"{total_pendente}")
    c5.metric("Sem Data Stábil", f"{total_sem_data}")
    c6.metric("Percentual On Time", texto_percentual,
              help=f"Sobre {ops_concluidas} OP(s) já entradas no CD. "
                   f"Pendentes ({total_pendente}) e sem data ({total_sem_data}) ficam fora do denominador.")

    st.caption(f"Base: {total_planilhadas} OP(s) | concluídas e classificáveis: {ops_concluidas}")

    st.write("---")
    df_em_transito = df_filtrado[df_filtrado['status_entrega'] == STATUS_PENDENTE].copy()
    total_pares_transito = df_em_transito['qtd_par_real_prensa'].sum()

    st.info(f"**⚠️ RESUMO DO TRÂNSITO (WIP):** Atualmente existem **{total_pendente}** ordens de produção "
            f"paradas na fábrica, totalizando **{total_pares_transito:,.0f}** pares prensados aguardando "
            "entrada física no CD.")

    aba_geral, aba_transito, aba_calendario = st.tabs([
        "📋 1. Base Integrada (Prensa + CD)",
        "🚨 2. Indicador Em Trânsito (WIP)",
        "📅 3. Calendário de Produção Prevista"
    ])

    with aba_geral:
        st.dataframe(
            df_filtrado[['op_acabado', 'cod_artigo', 'desc_artigo', 'data_stabil', 'data_limite_ontime',
                         'data_lanca_cd', 'qtd_par_real_prensa', 'status_entrega']],
            use_container_width=True
        )

    with aba_transito:
        if len(df_em_transito) > 0:
            st.dataframe(
                df_em_transito[['op_acabado', 'cod_artigo', 'desc_artigo', 'data_stabil',
                                'data_limite_ontime', 'qtd_par_real_prensa', 'status_entrega']],
                use_container_width=True
            )
        else:
            st.success("Não existem ordens de produção em trânsito para os critérios selecionados.")

    with aba_calendario:
        if df_calendario is not None:
            st.dataframe(df_calendario, use_container_width=True)
        else:
            st.warning("Calendário não pôde ser carregado.")
else:
    st.warning("Verifique os caminhos dos arquivos Excel em rede.")
