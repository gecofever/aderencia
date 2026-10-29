import pandas as pd
import datetime
import os
import streamlit as st

# Configuração da página do Streamlit
st.set_page_config(page_title="Dashboard PCP - On Time", layout="wide")

# Configuração visual do Pandas para não exibir notação científica
pd.set_option('display.float_format', lambda x: '%.2f' % x)

# ==============================================================================
# 1. CONFIGURAÇÕES DE CAMINHOS DE ARQUIVOS
# ==============================================================================
PATH_SAIDA_PRENSAS = r"M:\CGE-Pcp\Astro\Análises Programação\1- Prensas\Controle Horários de prensa\Saída de prensas.xlsx"
PATH_YMB51 = r"M:\CGE-Pcp\Pilar Planejamento\06 - Indicadores\3. Base de Dados\Entrada de CD\10.Outubro.XLSX"
PATH_CALENDARIO = r"M:\CGE-Pcp\Astro\Análises Programação\1- Prensas\Controle Horários de prensa\Saída de prensas.xlsx"

COLUNAS_ALVO = ['OP Acabado', 'Data Stábil', 'Qtd. Par Real', 'Cód. Artigo', 'Desc. Artigo']
COL_YMB51 = ['Ordem', 'Material', 'Qtd.  UM registro', 'Data de lançamento']
COL_CALENDARIO = ['Data', '6x1'] 

# ==============================================================================
# 2. FUNÇÕES DE PROCESSAMENTO DE DADOS
# ==============================================================================

@st.cache_data(ttl=600)
def carregar_calendario_previsto():
    try:
        df_cal = pd.read_excel(PATH_CALENDARIO, sheet_name='Balanceado', 
                               usecols=COL_CALENDARIO, engine='openpyxl', header=2)
        df_cal = df_cal.rename(columns={'Data': 'data_prevista', '6x1': 'status_6x1'})
        df_cal['data_prevista'] = pd.to_datetime(df_cal['data_prevista'], errors='coerce')
        df_cal = df_cal.dropna(subset=['data_prevista'])
        return df_cal
    except Exception as e:
        st.error(f"Erro ao ler a planilha de calendário: {e}")
        return None

def calcular_prazo_operacional(data_stabil, df_cal):
    if pd.isna(data_stabil) or df_cal is None:
        return pd.NaT
    data_limite = data_stabil + pd.Timedelta(days=3)
    datas_sem_producao = set(df_cal[df_cal['status_6x1'] == 0]['data_prevista'])
    while data_limite in datas_sem_producao:
        data_limite += pd.Timedelta(days=1)
    return data_limite

@st.cache_data(ttl=600)
def processar_base_integrada():
    try:
        df_p1 = pd.read_excel(PATH_SAIDA_PRENSAS, sheet_name='Stábil UGB 1', usecols=COLUNAS_ALVO, engine='openpyxl', header=2)
        df_p2 = pd.read_excel(PATH_SAIDA_PRENSAS, sheet_name='Stábil UGB 2', usecols=COLUNAS_ALVO, engine='openpyxl', header=2)
        df_acab1 = pd.read_excel(PATH_YMB51, sheet_name='Sheet1', usecols=COL_YMB51, engine='openpyxl', header=0)

        df_p1 = tratar_nomes(df_p1)
        df_p2 = tratar_nomes(df_p2)
        df_acab1 = tratar_nomes(df_acab1)

        if df_p1 is None or df_p2 is None or df_acab1 is None:
            return None

        # 1. CONSOLIDAÇÃO PRENSAS
        df_prensa_final = pd.concat([df_p1, df_p2], ignore_index=True).dropna(subset=['op_acabado'])
        df_prensa_final['op_acabado'] = df_prensa_final['op_acabado'].astype('Int64')
        df_prensa_final['cod_artigo'] = df_prensa_final['cod_artigo'].astype('Int64')
        df_prensa_final['data_stabil'] = pd.to_datetime(df_prensa_final['data_stabil'], errors='coerce')
        
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
        df_final = pd.merge(df_prensas_consolidado, df_acab_consolidado, on='op_acabado', how='left', suffixes=('_prensa', '_acabamento'))

        # 4. APLICAÇÃO DO STATUS ON TIME
        df_cal = carregar_calendario_previsto()
        if df_cal is not None:
            df_final['data_limite_ontime'] = df_final['data_stabil'].apply(lambda x: calcular_prazo_operacional(x, df_cal))
            
            def aplicar_status(row):
                if pd.isna(row['data_lanca_cd']): return 'Pendente'
                return 'On Time' if row['data_lanca_cd'] <= row['data_limite_ontime'] else 'Atrasado'
                    
            df_final['status_entrega'] = df_final.apply(aplicar_status, axis=1)
        else:
            df_final['data_limite_ontime'] = pd.NaT
            df_final['status_entrega'] = 'Sem Calendário'

        return df_final
    except Exception as e:
        st.error(f"Erro ao processar as bases: {e}")
        return None

def tratar_nomes(df):
    if df is None: return None
    renames_prensas = {'OP Acabado': 'op_acabado', 'Data Stábil': 'data_stabil', 'Qtd. Par Real': 'qtd_par_real', 'Cód. Artigo': 'cod_artigo', 'Desc. Artigo': 'desc_artigo'}
    renames_acab1 = {'Ordem': 'op_acabado', 'Material': 'cod_artigo', 'Qtd.  UM registro': 'qtd_par_real', 'Data de lançamento': 'data_lanca_cd'}
    if 'OP Acabado' in df.columns: return df.rename(columns=renames_prensas)
    elif 'Ordem' in df.columns: return df.rename(columns=renames_acab1)
    return df

# ==============================================================================
# 3. INTERFACE INTERATIVA (STREAMLIT)
# ==============================================================================

st.title("📊 Painel de Controle de Production - Indicador On Time")

df_base = processar_base_integrada()
df_calendario = carregar_calendario_previsto()

if df_base is not None:
    # --- FILTRO POR DIGITAÇÃO DE ARTIGO (Sidebar / Lateral) ---
    st.sidebar.header("🔍 Buscar por Artigo")
    
    # Campo de entrada numérica. Se deixar 0, assume a visão geral de todos
    busca_artigo = st.sidebar.number_input("Digite o Código do Artigo (ou 0 para todos):", min_value=0, value=0, step=1)
    
    if busca_artigo != 0:
        df_filtrado = df_base[df_base['cod_artigo'] == busca_artigo].copy()
        st.subheader(f"🎯 Resultados para o Artigo: {busca_artigo}")
    else:
        df_filtrado = df_base.copy()
        st.subheader("🌐 Visão Geral da Fábrica (Todos os Artigos)")

    # --- CÁLCULO DOS INDICADORES DOS CARDS ---
    total_planilhadas = len(df_filtrado)
    total_on_time = len(df_filtrado[df_filtrado['status_entrega'] == 'On Time'])
    total_atrasado = len(df_filtrado[df_filtrado['status_entrega'] == 'Atrasado'])
    total_pendente = len(df_filtrado[df_filtrado['status_entrega'] == 'Pendente'])
    
    # On Time % calculado de forma padrão sobre as OPs que já concluíram o processo (On Time + Atrasado)
    ops_concluidas = total_on_time + total_atrasado
    percentual_on_time = (total_on_time / ops_concluidas * 100) if ops_concluidas > 0 else 0.0

    # --- EXIBIÇÃO DE CARTÕES DE MÉTRICAS (KPIs) ---
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total de OPs Planilhadas", f"{total_planilhadas}")
    c2.metric("On Time", f"{total_on_time}", delta=None)
    c3.metric("Atrasado", f"{total_atrasado}", delta=None, delta_color="inverse")
    c4.metric("Pendente", f"{total_pendente}")
    c5.metric("Percentual On Time", f"{percentual_on_time:.1f}%")

    # --- INDICADOR ADICIONAL: RESUMO DO TRÂNSITO ---
    st.write("---")
    df_em_transito = df_filtrado[df_filtrado['status_entrega'] == 'Pendente'].copy()
    df_em_transito['qtd_par_real_acabamento'] = 0  # Ajuste conforme solicitado em etapas anteriores
    total_pares_transito = df_em_transito['qtd_par_real_prensa'].sum()

    st.info(f"**⚠️ RESUMO DO TRÂNSITO (WIP):** Atualmente existem **{total_pendente}** ordens de produção paradas na fábrica, totalizando **{total_pares_transito:,.0f}** pares prensados aguardando entrada física no CD.")

    # --- SEÇÃO DE ABAS DETALHADAS ---
    aba_geral, aba_transito, aba_calendario = st.tabs([
        "📋 1. Base Integrada (Prensa + CD)", 
        "🚨 2. Indicador Em Trânsito (WIP)", 
        "📅 3. Calendário de Produção Prevista"
    ])
    
    with aba_geral:
        st.dataframe(df_filtrado[['op_acabado', 'cod_artigo', 'desc_artigo', 'data_stabil', 'data_limite_ontime', 'data_lanca_cd', 'qtd_par_real_prensa', 'status_entrega']], use_container_width=True)
        
    with aba_transito:
        if len(df_em_transito) > 0:
            st.dataframe(df_em_transito[['op_acabado', 'cod_artigo', 'desc_artigo', 'data_stabil', 'data_limite_ontime', 'qtd_par_real_prensa', 'status_entrega']], use_container_width=True)
        else:
            st.success("Não existem ordens de produção em trânsito para os critérios selecionados.")
            
    with aba_calendario:
        if df_calendario is not None:
            st.dataframe(df_calendario, use_container_width=True)
        else:
            st.warning("Calendário não pôde ser carregado.")
else:
    st.warning("Verifique os caminhos dos arquivos Excel em rede.")