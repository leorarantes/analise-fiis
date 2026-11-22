# Coletor de indicadores de FIIs

Gera uma planilha Excel com indicadores de 25 fundos de investimento imobiliário (FIIs) da B3. O coletor usa o histórico diário de cotações, proventos e volume do Yahoo Finance, consultado via `yfinance`.

## Instalação e execução

Você precisa do pacote Python 3.10 ou superior instalado na sua máquina e conexão com a internet para consultar os dados. No terminal, dentro da pasta `coletor-fiis`, execute:

```bash
sudo apt install python3-venv
source .venv/bin/activate
pip install yfinance pandas openpyxl
python main.py
```

## O que a planilha contém

O arquivo `indicadores_fiis.xlsx` é salvo na mesma pasta de `main.py`. A aba **Indicadores** tem uma linha por fundo, e a aba **Notas** registra a fonte, a data da coleta e os critérios de cálculo.

| Indicador | Como é obtido |
| --- | --- |
| Rentabilidade acumulada | Retorno total em 2, 5 e 10 anos, com reinvestimento dos proventos. |
| Dividend yield (DY) atual | Proventos dos últimos 12 meses divididos pela cotação de fechamento mais recente. |
| DY médio de 5 anos | Média de cinco medições anuais do DY de 12 meses. |
| Liquidez média diária | Média de fechamento × volume nos últimos 12 meses, em reais. |
| Vacância física, vacância financeira e número de cotistas | Preenchimento manual ou importação pelo CSV opcional. |

Os indicadores que exigem um histórico maior do que o disponível ficam em branco. Se a coleta de um fundo falhar, seu ticker continua na planilha e o erro aparece no terminal. Para coletar outros FIIs, edite a lista `FIIS` em `main.py` e informe os códigos sem o sufixo `.SA`.

**Cada execução sobrescreve a planilha anterior.** Se você preencher células diretamente no Excel, guarde uma cópia ou registre esses dados no CSV antes de executar novamente.

## Como incluir vacância e número de cotistas

O coletor deixa essas colunas em amarelo. Você pode preenchê-las no Excel com dados de relatórios gerenciais, Fundamentus ou Status Invest. Para reaproveitar os dados a cada execução, crie `fundamentos_extras.csv` na mesma pasta de `main.py`:

```csv
Ticker,Vacância física (%),Vacância financeira (%),Cotistas
HGLG11,2.5,1.8,350000
```

Os valores da linha são ilustrativos. Use uma linha por ticker, informe a vacância em pontos percentuais (`2.5` representa 2,5%) e escreva o número de cotistas sem separador de milhar. Execute `python main.py` para incorporar os dados na planilha.
