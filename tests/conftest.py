import importlib.util
import os
import sys

import pandas as pd
import pytest

BASE = "/home/hermes/.hermes/cache/scratch"
FIXTURES = os.path.join(BASE, "aderencia_local")
MODULE_PATH = os.path.join(BASE, "aderencia_p0", "acabamento_stream.py")

# streamlit stub (não é um servidor Streamlit; só permite importar o dashboard)
sys.path.insert(0, os.path.join(BASE, "fake_streamlit"))


def carregar_modulo():
    spec = importlib.util.spec_from_file_location("acabamento_stream_p0", MODULE_PATH)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="session")
def mod():
    return carregar_modulo()


@pytest.fixture(scope="session")
def bases():
    """(df_base, df_calendario) gerados a partir das planilhas sintéticas."""
    if not os.path.exists(os.path.join(FIXTURES, "Saida de prensas.xlsx")):
        sys.path.insert(0, FIXTURES)
        import gerar_fixtures
        gerar_fixtures.planilha_prensas(os.path.join(FIXTURES, "Saida de prensas.xlsx"))
        gerar_fixtures.planilha_ymbt51(os.path.join(FIXTURES, "10.outubro.xlsx"))
    m = carregar_modulo()
    df = m.processar_base_integrada(
        path_prensas=os.path.join(FIXTURES, "Saida de prensas.xlsx"),
        path_cd=os.path.join(FIXTURES, "10.outubro.xlsx"),
        path_calendario=os.path.join(FIXTURES, "Saida de prensas.xlsx"),
    )
    cal = m.carregar_calendario_previsto(os.path.join(FIXTURES, "Saida de prensas.xlsx"))
    assert df is not None and cal is not None
    return df, cal


def status_de(df, op):
    return df.loc[df["op_acabado"] == op, "status_entrega"].iloc[0]


def prazo_de(df, op):
    return df.loc[df["op_acabado"] == op, "data_limite_ontime"].iloc[0]
