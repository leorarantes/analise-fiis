from argparse import ArgumentParser
from datetime import date, timedelta
from itertools import combinations
from pathlib import Path

import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FIIS = [
    # Logística
    "HGLG11", "XPLG11", "BTLG11", "BRCO11", "LVBI11",
    # Shoppings (MALL11 virou PMLL11 em julho de 2025)
    "HGBS11", "XPML11", "VISC11", "HSML11", "PMLL11",
    # Recebíveis (CRI)
    "MXRF11", "KNCR11", "KNIP11", "RBRR11", "HGCR11",
    # Lajes corporativas
    "PVBI11", "HGRE11", "JSRE11", "BRCR11", "TEPP11",
    # Híbridos / diversificados
    "KNRI11", "ALZR11", "HGRU11", "TRXF11", "TGAR11",
]

ARQUIVO_SAIDA = Path(__file__).with_name("indicadores_fiis.xlsx")
ARQUIVO_EXTRAS = Path(__file__).with_name("fundamentos_extras.csv")
TOLERANCIA_HISTORICO = pd.Timedelta(days=15)  # folga para fins de semana/feriados
LIMIAR_GRAFO = 0.60
INDICADORES_GRAFO = [
    "Rentabilidade acumulada 2 anos",
    "DY médio 2 anos",
    "Liquidez média diária 2 anos (R$)",
]

COLUNAS = [
    "Ticker",
    "Rentabilidade acumulada 2 anos",
    "Rentabilidade acumulada 5 anos",
    "Rentabilidade acumulada 10 anos",
    "DY atual (12m)",
    "DY médio 5 anos",
    "DY médio 2 anos",
    "Vacância física",
    "Vacância financeira",
    "Número de cotistas",
    "Liquidez média diária (R$)",
    "Liquidez média diária 2 anos (R$)",
]
COLUNAS_MANUAIS = ["Vacância física", "Vacância financeira", "Número de cotistas"]


# --------------------------------------------------------------------------
# Cálculo dos indicadores a partir do histórico (Close, Dividends, Volume)
# --------------------------------------------------------------------------
def tem_historico(df: pd.DataFrame, anos: int) -> bool:
    inicio = df.index[-1] - pd.DateOffset(years=anos)
    return df.index[0] <= inicio + TOLERANCIA_HISTORICO


def rentabilidade_acumulada(df: pd.DataFrame, anos: int):
    """Retorno total com reinvestimento dos proventos na data ex.

    r_t = (Close_t + Dividendo_t) / Close_{t-1} - 1; acumulado ao longo da janela.
    Devolve fração (0.35 = 35%) ou None se o fundo não tem histórico suficiente.
    """
    if not tem_historico(df, anos):
        return None
    inicio = df.index[-1] - pd.DateOffset(years=anos)
    janela = df.loc[df.index >= inicio]
    if len(janela) < 2:
        return None
    return retorno_total(janela)


def retorno_total(df: pd.DataFrame) -> float:
    close = df["Close"]
    prov = df["Dividends"].fillna(0.0)
    retornos = (close + prov) / close.shift(1) - 1
    return float((1 + retornos.iloc[1:]).prod() - 1)


def dy_janela(df: pd.DataFrame, fim: pd.Timestamp):
    """Proventos dos 12 meses até `fim` divididos pela cotação em `fim`."""
    ate_fim = df.loc[df.index <= fim]
    if ate_fim.empty:
        return None
    preco = float(ate_fim["Close"].iloc[-1])
    prov = float(
        df.loc[(df.index > fim - pd.DateOffset(years=1)) & (df.index <= fim), "Dividends"].sum()
    )
    return prov / preco if preco else None


def dy_medio(df: pd.DataFrame, anos: int = 5):
    """Média dos DY de 12 meses medidos ao fim de cada um dos últimos `anos` anos."""
    if not tem_historico(df, anos):
        return None
    fim = df.index[-1]
    valores = [dy_janela(df, fim - pd.DateOffset(years=k)) for k in range(anos)]
    valores = [v for v in valores if v is not None]
    return sum(valores) / len(valores) if valores else None


def liquidez_media_diaria(df: pd.DataFrame):
    """Média diária de Close x Volume (R$) nos últimos 12 meses."""
    janela = df.loc[df.index > df.index[-1] - pd.DateOffset(years=1)]
    if janela.empty or "Volume" not in janela:
        return None
    return float((janela["Close"] * janela["Volume"]).mean())


def indicadores_do_grafo(df: pd.DataFrame, corte: pd.Timestamp) -> dict:
    inicio = corte - pd.DateOffset(years=2)
    janela = df.loc[(df.index > inicio) & (df.index <= corte)]
    base = df.loc[df.index <= inicio].tail(1)
    if janela.empty or base.empty or inicio - base.index[0] > TOLERANCIA_HISTORICO:
        return {}
    precos = pd.concat([base, janela])
    if (
        corte - janela.index[-1] > TOLERANCIA_HISTORICO
        or precos.index.to_series().diff().gt(TOLERANCIA_HISTORICO).any()
    ):
        return {}

    indicadores = {}
    if "Volume" in janela and janela["Volume"].notna().all():
        indicadores[INDICADORES_GRAFO[2]] = float((janela["Close"] * janela["Volume"]).mean())

    if "Dividends" not in df:
        return indicadores
    if janela["Dividends"].notna().all():
        indicadores[INDICADORES_GRAFO[0]] = retorno_total(precos)

    mensal = []
    ultimo_mes = corte.to_period("M")
    if ultimo_mes.end_time.normalize() > corte:
        ultimo_mes -= 1
    for mes in pd.period_range(end=ultimo_mes, periods=24, freq="M"):
        fim = mes.end_time.normalize()
        inicio_dy = fim - pd.DateOffset(years=1)
        historico = df.loc[(df.index > inicio_dy) & (df.index <= fim)]
        if (
            df.index[0] > inicio_dy
            or historico.empty
            or fim - historico.index[-1] > TOLERANCIA_HISTORICO
            or historico.index[0] - inicio_dy > TOLERANCIA_HISTORICO
            or historico.index.to_series().diff().gt(TOLERANCIA_HISTORICO).any()
            or historico["Dividends"].isna().any()
        ):
            break
        mensal.append(dy_janela(historico, fim))
    if len(mensal) == 24 and all(valor is not None for valor in mensal):
        indicadores[INDICADORES_GRAFO[1]] = sum(mensal) / len(mensal)
    return indicadores


# --------------------------------------------------------------------------
# Coleta
# --------------------------------------------------------------------------
def coletar(ticker: str, corte: date) -> dict:
    df = yf.Ticker(f"{ticker}.SA").history(
        period="max", interval="1d", auto_adjust=False, actions=True
    )
    if df.empty or "Close" not in df:
        raise ValueError("histórico vazio")
    df.index = df.index.tz_localize(None).normalize()
    df = df.sort_index()
    df = df.loc[~df.index.duplicated(keep="last")]
    data_corte = pd.Timestamp(corte)
    df = df.loc[(df.index <= data_corte) & (df["Close"] > 0)]
    if df.empty:
        raise ValueError("sem cotações válidas até a data de corte")
    perfil = indicadores_do_grafo(df, data_corte)
    atual = df.loc[df.index > data_corte - pd.DateOffset(years=1)]
    dy_descritivo = (
        dy_janela(df, data_corte)
        if "Dividends" in atual
        and atual["Dividends"].notna().all()
        and df.index[0] <= data_corte - pd.DateOffset(years=1)
        and data_corte - df.index[-1] <= TOLERANCIA_HISTORICO
        else None
    )
    if "Dividends" not in df:
        df["Dividends"] = 0.0

    return {
        "Ticker": ticker,
        "Rentabilidade acumulada 5 anos": rentabilidade_acumulada(df, 5),
        "Rentabilidade acumulada 10 anos": rentabilidade_acumulada(df, 10),
        "DY atual (12m)": dy_descritivo,
        "DY médio 5 anos": dy_medio(df, 5),
        "Liquidez média diária (R$)": liquidez_media_diaria(df),
        **perfil,
    }


def ler_extras() -> pd.DataFrame | None:
    if not ARQUIVO_EXTRAS.exists():
        return None
    bruto = pd.read_csv(ARQUIVO_EXTRAS, sep=None, engine="python", dtype=str)
    bruto["Ticker"] = bruto["Ticker"].str.strip().str.upper().str.removesuffix(".SA")

    def numero(coluna):
        return pd.to_numeric(
            bruto[coluna].str.replace("%", "", regex=False).str.replace(",", ".", regex=False),
            errors="coerce",
        )

    extras = pd.DataFrame({"Ticker": bruto["Ticker"]})
    for col in ("Vacância física", "Vacância financeira"):
        if f"{col} (%)" in bruto:
            extras[col] = numero(f"{col} (%)") / 100  # pontos -> fração
    if "Cotistas" in bruto:
        extras["Número de cotistas"] = numero("Cotistas")
    return extras.set_index("Ticker")


# --------------------------------------------------------------------------
# Planilha
# --------------------------------------------------------------------------
def calcular_grafo(linhas: list[dict]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, list]:
    indicadores = pd.DataFrame(linhas).set_index("Ticker").reindex(columns=INDICADORES_GRAFO)
    indicadores = indicadores.astype(float).replace([float("inf"), -float("inf")], float("nan"))
    completos = indicadores.dropna()
    desvios = completos.std(ddof=0).where(completos.nunique().gt(1))
    escores = ((completos - completos.mean()) / desvios).reindex(indicadores.index)
    correlacoes = escores.T.corr(min_periods=len(INDICADORES_GRAFO))
    # O pandas fixa a diagonal em 1 mesmo para perfis constantes.
    validos = escores.notna().all(axis=1) & escores.nunique(axis=1).gt(1)
    correlacoes.loc[~validos, :] = float("nan")
    correlacoes.loc[:, ~validos] = float("nan")
    pesos = (correlacoes + 1) / 2

    pares = []
    for fii1, fii2 in combinations(indicadores.index, 2):
        rho = correlacoes.at[fii1, fii2]
        peso = pesos.at[fii1, fii2]
        if pd.isna(peso):
            motivo = (
                "Indicadores ausentes"
                if indicadores.loc[[fii1, fii2]].isna().any().any()
                else "Correlação indefinida (variância zero)"
            )
            pares.append((fii1, fii2, None, None, False, LIMIAR_GRAFO, motivo))
        else:
            incluir = peso >= LIMIAR_GRAFO
            pares.append((fii1, fii2, float(rho), float(peso), bool(incluir), LIMIAR_GRAFO,
                          "Aresta" if incluir else "Abaixo do limiar"))
    pares.sort(key=lambda par: par[3] if par[3] is not None else -1, reverse=True)
    return escores, correlacoes, pesos, pares


def gravar_abas_grafo(
    wb: Workbook, linhas: list[dict], fonte: Font, negrito: Font, cabecalho_fill: PatternFill
) -> tuple[int, int, int]:
    escores, correlacoes, pesos, pares = calcular_grafo(linhas)
    for nome, tabela in [("Z_scores", escores), ("Correlacao", correlacoes), ("Pesos_W", pesos)]:
        ws = wb.create_sheet(nome)
        ws.append(["Ticker", *tabela.columns])
        for ticker, *valores in tabela.itertuples(name=None):
            ws.append([ticker, *(None if pd.isna(valor) else float(valor) for valor in valores)])
        ws.freeze_panes = "B2"
        for coluna in ws.iter_cols(min_col=2, min_row=2):
            for celula in coluna:
                celula.number_format = "0.0000"
        if nome == "Z_scores":
            for indice in range(2, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(indice)].width = 32

    ws = wb.create_sheet("Pares_do_grafo")
    ws.append(["FII 1", "FII 2", "Correlação (rho)", "Peso W (0 a 1)",
               "Incluir com limiar atual", "Limiar usado", "Situação"])
    for par in pares:
        ws.append(par)
        for celula in ws[ws.max_row][2:4]:
            celula.number_format = "0.0000"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for coluna, largura in zip("CDEFG", [18, 18, 25, 16, 42]):
        ws.column_dimensions[coluna].width = largura

    for nome in ("Z_scores", "Correlacao", "Pesos_W", "Pares_do_grafo"):
        ws = wb[nome]
        for celula in ws[1]:
            celula.font = negrito
            celula.fill = cabecalho_fill
            celula.alignment = Alignment(wrap_text=True)
        for linha in ws.iter_rows(min_row=2):
            for celula in linha:
                celula.font = fonte
        ws.row_dimensions[1].height = 32
    vertices = int(correlacoes.notna().any().sum())
    return vertices, sum(par[3] is not None for par in pares), sum(par[4] for par in pares)


def gravar_xlsx(linhas: list[dict], corte: date) -> None:
    fonte = Font(name="Arial", size=10)
    negrito = Font(name="Arial", size=10, bold=True)
    cabecalho_fill = PatternFill("solid", fgColor="D9E1F2")
    manual_fill = PatternFill("solid", fgColor="FFF2CC")

    formatos = {
        "Rentabilidade acumulada 2 anos": "0.00%",
        "Rentabilidade acumulada 5 anos": "0.00%",
        "Rentabilidade acumulada 10 anos": "0.00%",
        "DY atual (12m)": "0.00%",
        "DY médio 5 anos": "0.00%",
        "DY médio 2 anos": "0.00%",
        "Vacância física": "0.00%",
        "Vacância financeira": "0.00%",
        "Número de cotistas": "#,##0",
        "Liquidez média diária (R$)": '"R$" #,##0',
        "Liquidez média diária 2 anos (R$)": '"R$" #,##0',
    }

    wb = Workbook()
    ws = wb.active
    ws.title = "Indicadores"
    for j, nome in enumerate(COLUNAS, 1):
        c = ws.cell(1, j, nome)
        c.font, c.fill = negrito, cabecalho_fill
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 42

    for i, linha in enumerate(linhas, 2):
        for j, nome in enumerate(COLUNAS, 1):
            valor = linha.get(nome)
            if valor is not None and pd.isna(valor):
                valor = None
            c = ws.cell(i, j, valor)
            c.font = fonte
            if nome in formatos:
                c.number_format = formatos[nome]
            if nome in COLUNAS_MANUAIS:
                c.fill = manual_fill

    for j, nome in enumerate(COLUNAS, 1):
        ws.column_dimensions[get_column_letter(j)].width = 10 if j == 1 else 17
    ws.freeze_panes = "B2"

    # Notas com definições, fontes e premissas
    nt = wb.create_sheet("Notas")
    vertices, pares_validos, arestas = gravar_abas_grafo(wb, linhas, fonte, negrito, cabecalho_fill)
    inicio = pd.Timestamp(corte) - pd.DateOffset(years=2)
    notas = [
        ("Fonte dos dados de mercado", "yfinance (Yahoo Finance), tickers com sufixo .SA, histórico diário máximo"),
        ("Data da coleta", date.today().strftime("%d/%m/%Y")),
        ("Janela comum do grafo", f"De {inicio:%d/%m/%Y} a {corte:%d/%m/%Y}. "
                                  "Tolera até 15 dias corridos sem cotação."),
        ("Rentabilidade acumulada", "Retorno total com reinvestimento dos proventos na data ex. "
                                    "Em 2 anos, usa a última cotação até o início da janela comum como base. "
                                    "As janelas descritivas de 5 e 10 anos terminam na última cotação do fundo."),
        ("DY atual (12m)", "Proventos dos últimos 12 meses / cotação de fechamento mais recente."),
        ("DY médio 5 anos", "Média de 5 DYs de 12 meses, medidos ao fim de cada um dos últimos 5 anos "
                            "(proventos de 12m / cotação no fim da janela). Em branco se houver menos de 5 anos de histórico."),
        ("DY médio 2 anos", "Média dos 24 DYs de 12 meses no final dos últimos 24 meses completos. "
                            "Exige cobertura para todas as medições."),
        ("Liquidez média diária (R$)", "Média de Close x Volume nos últimos 12 meses."),
        ("Liquidez média diária 2 anos", "Média de Close x Volume na janela comum, sem transformação logarítmica."),
        ("Padronização e peso", "Escore z com desvio padrão populacional entre fundos com os três indicadores. "
                                "Pearson entre os dois perfis completos. Peso W = (rho + 1) / 2."),
        ("Grafo de similaridade", f"Limiar W ≥ {LIMIAR_GRAFO:.2f}. {vertices} de {len(linhas)} FIIs com "
                                   f"perfil válido, {pares_validos} pares com peso e {arestas} arestas."),
        ("Vacância física, vacância financeira e número de cotistas",
         "Não disponíveis no yfinance. Células amarelas para preenchimento manual (Fundamentus, StatusInvest "
         "ou relatório gerencial) ou via fundamentos_extras.csv. Digite vacância como percentual (ex.: 5,2%)."),
    ]
    for i, (titulo, texto) in enumerate(notas, 1):
        a, b = nt.cell(i, 1, titulo), nt.cell(i, 2, texto)
        a.font, b.font = negrito, fonte
        a.alignment = Alignment(wrap_text=True, vertical="top")
        b.alignment = Alignment(wrap_text=True, vertical="top")
    nt.column_dimensions["A"].width = 34
    nt.column_dimensions["B"].width = 100

    wb.save(ARQUIVO_SAIDA)


def main() -> None:
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--data-corte", type=date.fromisoformat,
                        default=date.today() - timedelta(days=1),
                        help="Data comum até a qual usar cotações (AAAA-MM-DD). Padrão: ontem.")
    corte = parser.parse_args().data_corte
    if corte >= date.today():
        parser.error("--data-corte deve ser anterior a hoje")
    linhas, falhas = [], []
    for ticker in FIIS:
        print(f"Coletando {ticker}...")
        try:
            linhas.append(coletar(ticker, corte))
        except Exception as exc:  # mantém o ticker na planilha, com indicadores vazios
            falhas.append((ticker, str(exc)))
            linhas.append({"Ticker": ticker})

    extras = ler_extras()
    if extras is not None:
        for linha in linhas:
            if linha["Ticker"] in extras.index:
                linha.update(extras.loc[linha["Ticker"]].to_dict())
        print(f"Dados extras incluídos de {ARQUIVO_EXTRAS.name}")

    gravar_xlsx(linhas, corte)
    print(f"\nPlanilha criada: {ARQUIVO_SAIDA}")
    print(f"FIIs coletados: {len(FIIS) - len(falhas)} de {len(FIIS)}")
    for ticker, erro in falhas:
        print(f"  falhou: {ticker} ({erro})")


if __name__ == "__main__":
    main()
