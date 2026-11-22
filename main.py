"""Monta uma planilha com os indicadores de FIIs da B3, via yfinance.

Indicadores (somente estes):
    - Rentabilidade acumulada (2, 5 e 10 anos)
    - Dividend Yield (DY) atual e média de 5 anos
    - Vacância física e vacância financeira
    - Número de cotistas
    - Liquidez média diária

O yfinance NÃO fornece vacância nem número de cotistas para FIIs. Essas três
colunas saem em branco (amarelas) para preenchimento manual com dados do
Fundamentus, StatusInvest ou dos relatórios gerenciais. Se existir um arquivo
fundamentos_extras.csv na mesma pasta, elas são preenchidas automaticamente:

    Ticker,Vacância física (%),Vacância financeira (%),Cotistas
    HGLG11,2.5,1.8,350000

    (percentuais em pontos, ex.: 2.5 = 2,5%; cotistas sem separador de milhar)

Uso:
    pip install yfinance pandas openpyxl
    python coletar_indicadores_fiis.py

Saída: indicadores_fiis.xlsx (mesma pasta do script).
"""

from datetime import date
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

COLUNAS = [
    "Ticker",
    "Rentabilidade acumulada 2 anos",
    "Rentabilidade acumulada 5 anos",
    "Rentabilidade acumulada 10 anos",
    "DY atual (12m)",
    "DY médio 5 anos",
    "Vacância física",
    "Vacância financeira",
    "Número de cotistas",
    "Liquidez média diária (R$)",
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
    close = janela["Close"]
    prov = janela["Dividends"].fillna(0.0)
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


def dy_atual(df: pd.DataFrame):
    return dy_janela(df, df.index[-1]) if tem_historico(df, 1) else None


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


# --------------------------------------------------------------------------
# Coleta
# --------------------------------------------------------------------------
def coletar(ticker: str) -> dict:
    df = yf.Ticker(f"{ticker}.SA").history(
        period="max", interval="1d", auto_adjust=False, actions=True
    )
    if df.empty or "Close" not in df:
        raise ValueError("histórico vazio")
    df = df.dropna(subset=["Close"]).sort_index()
    if "Dividends" not in df:
        df["Dividends"] = 0.0

    return {
        "Ticker": ticker,
        "Rentabilidade acumulada 2 anos": rentabilidade_acumulada(df, 2),
        "Rentabilidade acumulada 5 anos": rentabilidade_acumulada(df, 5),
        "Rentabilidade acumulada 10 anos": rentabilidade_acumulada(df, 10),
        "DY atual (12m)": dy_atual(df),
        "DY médio 5 anos": dy_medio(df, 5),
        "Liquidez média diária (R$)": liquidez_media_diaria(df),
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
def gravar_xlsx(linhas: list[dict]) -> None:
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
        "Vacância física": "0.00%",
        "Vacância financeira": "0.00%",
        "Número de cotistas": "#,##0",
        "Liquidez média diária (R$)": '"R$" #,##0',
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
    notas = [
        ("Fonte dos dados de mercado", "yfinance (Yahoo Finance), tickers com sufixo .SA, histórico diário máximo"),
        ("Data da coleta", date.today().strftime("%d/%m/%Y")),
        ("Rentabilidade acumulada", "Retorno total com reinvestimento dos proventos na data ex: "
                                    "(Close + Dividendo) / Close anterior, acumulado na janela de 2, 5 e 10 anos "
                                    "até o último pregão. Em branco se o fundo tem histórico menor que a janela."),
        ("DY atual (12m)", "Proventos dos últimos 12 meses / cotação de fechamento mais recente."),
        ("DY médio 5 anos", "Média de 5 DYs de 12 meses, medidos ao fim de cada um dos últimos 5 anos "
                            "(proventos de 12m / cotação no fim da janela). Em branco se houver menos de 5 anos de histórico."),
        ("Liquidez média diária (R$)", "Média de Close x Volume nos últimos 12 meses."),
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
    linhas, falhas = [], []
    for ticker in FIIS:
        print(f"Coletando {ticker}...")
        try:
            linhas.append(coletar(ticker))
        except Exception as exc:  # mantém o ticker na planilha, com indicadores vazios
            falhas.append((ticker, str(exc)))
            linhas.append({"Ticker": ticker})

    extras = ler_extras()
    if extras is not None:
        for linha in linhas:
            if linha["Ticker"] in extras.index:
                linha.update(extras.loc[linha["Ticker"]].to_dict())
        print(f"Dados extras incluídos de {ARQUIVO_EXTRAS.name}")

    gravar_xlsx(linhas)
    print(f"\nPlanilha criada: {ARQUIVO_SAIDA}")
    print(f"FIIs coletados: {len(FIIS) - len(falhas)} de {len(FIIS)}")
    for ticker, erro in falhas:
        print(f"  falhou: {ticker} ({erro})")


if __name__ == "__main__":
    main()