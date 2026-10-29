"""Testes de regressão do P0 do indicador On Time (PCP)."""
import os

import pandas as pd
import pytest

BASE = "/home/hermes/.hermes/cache/scratch"
FIXTURES = os.path.join(BASE, "aderencia_local")


# ---------------------------------------------------------------- P0.1 (hora)
def test_lancamento_no_mesmo_dia_do_prazo_e_on_time(bases):
    """Lançamento no dia limite, com hora > 00:00, não pode virar 'Atrasado'."""
    df, _ = bases
    assert pd.Timestamp(df.loc[df.op_acabado == 1006, "data_lanca_cd"].iloc[0]) == pd.Timestamp("2026-10-09 09:00")
    assert df.loc[df.op_acabado == 1006, "status_entrega"].iloc[0] == "On Time"


def test_lancamento_tarde_da_noite_no_dia_limite_e_on_time(bases):
    df, _ = bases
    assert df.loc[df.op_acabado == 2001, "data_lanca_cd"].iloc[0] == pd.Timestamp("2026-10-13 23:00")
    assert df.loc[df.op_acabado == 2001, "status_entrega"].iloc[0] == "On Time"


def test_lancamento_apos_o_prazo_continua_atrasado(mod):
    """Garante que a normalização não mascarou atraso real."""
    cal = pd.DataFrame({"data_prevista": pd.to_datetime(["2026-10-01", "2026-10-02"]),
                        "status_6x1": [1, 1]})
    df = pd.DataFrame({
        "data_stabil": pd.to_datetime(["2026-10-01"]),
        "data_lanca_cd": pd.to_datetime(["2026-10-05 10:00"]),
    })
    df["data_limite_ontime"] = mod.calcular_prazo_operacional(df["data_stabil"], cal)
    assert mod.classificar_status(df).iloc[0] == "Atrasado"


# ------------------------------------------------------------- P0.2 (sem data)
def test_op_sem_data_stabil_nao_e_atrasada(bases):
    df, _ = bases
    assert df.loc[df.op_acabado == 1007, "status_entrega"].iloc[0] == "Sem Data"


def test_op_sem_data_nao_entra_no_denominador(bases):
    df, _ = bases
    contagem = df["status_entrega"].value_counts()
    # 1002 e 1007 não têm Data Stábil -> 'Sem Data' (fora do denominador)
    assert int(contagem.get("Sem Data", 0)) == 2
    assert int(contagem.get("On Time", 0)) == 4
    assert int(contagem.get("Atrasado", 0)) == 0
    assert int(contagem.get("Pendente", 0)) == 2
    concluidas = int(contagem.get("On Time", 0)) + int(contagem.get("Atrasado", 0))
    assert concluidas == 4
    assert int(contagem.get("On Time", 0)) / concluidas * 100 == 100.0


# --------------------------------------------------- P0.3 (calendário e texto)
def test_calendario_com_6x1_em_texto_e_convertido(mod, tmp_path):
    caminho = tmp_path / "calendario_texto.xlsx"
    linhas = [["cab", "calho"], ["lixo", "lixo"], ["Data", "6x1"],
              ["2026-10-01", "1"], ["2026-10-02", "0"], ["2026-10-03", "0"]]
    with pd.ExcelWriter(caminho, engine="openpyxl") as xl:
        pd.DataFrame(linhas).to_excel(xl, sheet_name="Balanceado", header=False, index=False)

    cal = mod.carregar_calendario_previsto(str(caminho))
    assert cal is not None
    assert pd.api.types.is_numeric_dtype(cal["status_6x1"])
    dias, inicios, fins = mod._blocos_sem_producao(cal)
    assert dias.size == 2
    assert fins[0] - inicios[0] == 1  # bloco contíguo de 2 dias


def test_prazo_desloca_bloco_contiguo_de_parada(bases):
    """Data Stábil 01/10 + 3 = 04/10, que cai em parada 04-06/10 => prazo 07/10."""
    df, _ = bases
    assert pd.Timestamp(df.loc[df.op_acabado == 1001, "data_limite_ontime"].iloc[0]) == pd.Timestamp("2026-10-07")


def test_prazo_de_op_em_dia_com_producao_nao_muda(mod):
    cal = pd.DataFrame({"data_prevista": pd.to_datetime(["2026-09-02", "2026-10-11"]),
                        "status_6x1": [0, 0]})
    df = pd.DataFrame({"data_stabil": pd.to_datetime(["2026-09-01"])})
    prazo = mod.calcular_prazo_operacional(df["data_stabil"], cal)
    assert prazo.iloc[0] == pd.Timestamp("2026-09-04")


def test_prazo_vetorizado_igual_ao_calculo_ingenuo(mod):
    """Compara a versão vetorizada com o while original em 2.000 datas aleatórias."""
    import numpy as np
    rng = np.random.default_rng(7)
    datas_sem_prod = np.sort(pd.to_datetime(
        ["2026-10-04", "2026-10-05", "2026-10-06", "2026-10-11", "2026-10-12", "2026-10-25"]))
    cal = pd.DataFrame({"data_prevista": datas_sem_prod, "status_6x1": 0})
    estabil = pd.Series(pd.to_datetime("2026-09-01") + pd.to_timedelta(rng.integers(0, 60, 2000), unit="D"))

    def ingenuo(data):
        if pd.isna(data):
            return pd.NaT
        limite = data + pd.Timedelta(days=3)
        paradas = set(datas_sem_prod)
        while limite in paradas:
            limite += pd.Timedelta(days=1)
        return limite

    esperado = estabil.apply(ingenuo)
    obtido = mod.calcular_prazo_operacional(estabil, cal)
    assert (esperado.values == obtido.values).all()


# ------------------------------------------------------- P0.4 (cobertura/limite)
def test_prazo_alem_do_calendario_e_sinalizado(bases):
    """1004 (08/10 + 3 = 11/10, parada 11-12 => 13/10) está no limite do calendário."""
    df, cal = bases
    ultimo = cal["data_prevista"].max()
    assert pd.Timestamp(df.loc[df.op_acabado == 1004, "data_limite_ontime"].iloc[0]) == pd.Timestamp("2026-10-13")
    assert ultimo == pd.Timestamp("2026-10-13")


# --------------------------------------------------------------- consistência
def test_todas_as_ops_estao_classificadas(bases):
    df, _ = bases
    assert df["status_entrega"].isin(["On Time", "Atrasado", "Pendente", "Sem Data"]).all()
    assert not df["status_entrega"].isna().any()


def test_quantidade_consolidada_por_op(bases):
    df, _ = bases
    assert int(df.loc[df.op_acabado == 1006, "qtd_par_real_prensa"].iloc[0]) == 650
