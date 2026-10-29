import pandas as pd
import datetime
import os
import streamlit as st

# Configuração da página do Streamlit (deve ser o primeiro comando Streamlit)
st.set_page_config(page_title="Dashboard de Aderência Prensas", layout="wide")

# 1. CONFIGURAÇÕES DE CAMINHOS
PATH_SAIDA_PRENSAS = r"M:\CGE-Pcp\Astro\Análises Programação\1- Prensas\Controle Horários de prensa\Saída de prensas.xlsx"
COLUNAS_ALVO = ['Data Stábil', 'Status', 'Qtd. Par Real', 'Cód. Artigo', 'Desc. Artigo']

@st.cache_data(ttl=600)  # Guarda os dados em cache por 10 minutos para carregar instantaneamente
def carregar_e_unificar_dados():
    try:
        # 1. Leitura das abas
        df_p1 = pd.read_excel(PATH_SAIDA_PRENSAS, sheet_name='Stábil UGB 1',
                              usecols=COLUNAS_ALVO, engine='openpyxl', header=2)
        df_p2 = pd.read_excel(PATH_SAIDA_PRENSAS, sheet_name='Stábil UGB 2',
                              usecols=COLUNAS_ALVO, engine='openpyxl', header=2)

        df_p1 = tratar_nomes(df_p1)
        df_p2 = tratar_nomes(df_p2)

        if df_p1 is None or df_p2 is None:
            return None

        # Identificadores de UGB
        df_p1['ugb'] = 'UGB 1'
        df_p2['ugb'] = 'UGB 2'

        # Consolidação completa da base
        df_prensa_final = pd.concat([df_p1, df_p2], ignore_index=True)
        return df_prensa_final
    
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        return None

def tratar_nomes(df):
    if df is None:
        return None
    renames = {
        'Data Stábil': 'data_stabil', 'Status': 'status', 'Qtd. Par Real': 'qtd_par_real',
        'Cód. Artigo': 'cod_artigo', 'Desc. Artigo': 'desc_artigo'
    }
    return df.rename(columns=renames)

def processar_base_indicadores(df):
    """Gera a tabela aberta por artigo e UGB para alimentar os filtros do Streamlit"""
    # 1. Total produzido por Artigo + UGB
    df_total = df.groupby(['cod_artigo', 'desc_artigo', 'ugb'])['qtd_par_real'].sum().reset_index()
    df_total = df_total.rename(columns={'qtd_par_real': 'qtd_total_produzida'})
    
    # 2. Total On Time por Artigo + UGB
    filtro_on_time = df[df['status'].str.strip() == 'On Time']
    df_on_time = filtro_on_time.groupby(['cod_artigo', 'ugb'])['qtd_par_real'].sum().reset_index()
    df_on_time = df_on_time.rename(columns={'qtd_par_real': 'qtd_on_time'})
    
    # 3. Cruzamento mantendo granularidade de UGB
    df_res = pd.merge(df_total, df_on_time, on=['cod_artigo', 'ugb'], how='left')
    df_res['qtd_on_time'] = df_res['qtd_on_time'].fillna(0)
    
    # 4. Cálculo percentual por linha
    df_res['perct_ontime'] = (df_res['qtd_on_time'] / df_res['qtd_total_produzida']) * 100
    df_res['perct_ontime'] = df_res['perct_ontime'].round(2)
    
    return df_res.sort_values(by='qtd_total_produzida', ascending=False)


# --- INTERFACE STREAMLIT ---

st.title("📊 Controle de Aderência e Horários de Prensa")
st.markdown("---")

# Carregamento de dados com indicador visual de progresso
with st.spinner("Buscando dados na rede PCP... Por favor, aguarde."):
    df_consolidado = carregar_e_unificar_dados()

if df_consolidado is not None:
    
    # --- CÁLCULO DOS CARDS GLOBAIS (Base Total Sem Filtros) ---
    total_geral_produzido = int(df_consolidado['qtd_par_real'].sum())
    
    filtro_ot_global = df_consolidado[df_consolidado['status'].str.strip() == 'On Time']
    total_geral_on_time = int(filtro_ot_global['qtd_par_real'].sum())
    
    on_time_geral_percentual = (total_geral_on_time / total_geral_produzido) * 100 if total_geral_produzido > 0 else 0

    # --- SEÇÃO DE CARDS ---
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(label="Total Geral Produzido (Base)", value=f"{total_geral_produzido:,}".replace(",", "."))
        
    with col2:
        st.metric(label="Total Geral On Time (Base)", value=f"{total_geral_on_time:,}".replace(",", "."))
        
    with col3:
        st.metric(label="Percentual Geral On Time", value=f"{on_time_geral_percentual:.2f} %")

    st.markdown("---")
    
    # --- SEÇÃO DE FILTROS DINÂMICOS ---
    st.subheader("🔍 Filtros de Consulta")
    
    # Processa o dataframe final estruturado por Artigo/UGB
    df_on_time_percentual = processar_base_indicadores(df_consolidado)
    
    c_filtro1, c_filtro2 = st.columns(2)
    
    with c_filtro1:
        # Filtro por UGB
        opcoes_ugb = ["Todas"] + list(df_on_time_percentual['ugb'].unique())
        ugb_selecionada = st.selectbox("Selecione a UGB:", opcoes_ugb)
        
    with c_filtro2:
        # Filtro de busca por Artigo (Código ou Descrição)
        pesquisa_artigo = st.text_input("Buscar por Código do Artigo (deixe vazio para todos):").strip()

    # Aplicação dinâmica dos filtros no DataFrame de exibição
    df_filtrado = df_on_time_percentual.copy()
    
    if ugb_selecionada != "Todas":
        df_filtrado = df_filtrado[df_filtrado['ugb'] == ugb_selecionada]
        
    if pesquisa_artigo:
        df_filtrado = df_filtrado[df_filtrado['cod_artigo'].astype(str).str.contains(pesquisa_artigo, case=False)]

    # --- TABELA DE RESULTADOS ---
    st.subheader("📋 Tabela de Resultados - Detalhado por Artigo")
    
    # Formatação visual das colunas numéricas para exibição no Streamlit
    df_formatado = df_filtrado.copy()
    df_formatado = df_formatado.rename(columns={
        'cod_artigo': 'Cód. Artigo',
        'desc_artigo': 'Descrição',
        'ugb': 'UGB',
        'qtd_total_produzida': 'Qtd. Total Produzida',
        'qtd_on_time': 'Qtd. On Time',
        'perct_ontime': '% On Time'
    })
    
    st.dataframe(
        df_formatado.style.format({
            'Qtd. Total Produzida': '{:,.0f}',
            'Qtd. On Time': '{:,.0f}',
            '% On Time': '{:.2f} %'
        }), 
        use_container_width=True,
        hide_index=True
    )
else:
    st.error("Não foi possível renderizar o dashboard devido a falhas no carregamento da base.")