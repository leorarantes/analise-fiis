# Coletor de indicadores de FIIs

Gera uma planilha Excel com indicadores e um grafo de similaridade de perfil para 25 fundos de investimento imobiliário (FIIs) da B3. O coletor usa o histórico diário de cotações, proventos e volume do Yahoo Finance, consultado via `yfinance`.

## Instalação e execução

Você precisa do pacote Python 3.10 ou superior instalado na sua máquina e conexão com a internet para consultar os dados. No terminal, dentro da pasta `coletor-fiis`, execute:

```bash
sudo apt install python3-venv
python3 -m venv .venv
source .venv/bin/activate
pip install yfinance pandas openpyxl
python main.py
```

O comando gera as abas de indicadores e do grafo na mesma planilha. Por padrão, a data de corte é ontem. Para usar a data de referência da metodologia, execute `python main.py --data-corte 2026-11-17`. A mesma data vale para todos os fundos e fica registrada na aba **Notas**.

## O que a planilha contém

O arquivo `indicadores_fiis.xlsx` é salvo na mesma pasta de `main.py`. A aba **Indicadores** tem uma linha por fundo, e a aba **Notas** registra a fonte, a data da coleta e os critérios de cálculo.

| Indicador | Como é obtido |
| --- | --- |
| Rentabilidade acumulada | Retorno total em 2, 5 e 10 anos, com reinvestimento dos proventos. |
| Dividend yield (DY) atual | Proventos dos últimos 12 meses divididos pela cotação de fechamento mais recente. |
| DY médio de 5 anos | Média de cinco medições anuais do DY de 12 meses. |
| DY médio de 2 anos | Média do DY de 12 meses medido no fim de cada um dos últimos 24 meses completos. |
| Liquidez média diária | Média de fechamento × volume nos últimos 12 meses, em reais. |
| Liquidez média diária de 2 anos | Média de fechamento × volume na janela comum do grafo, em reais. |
| Vacância física, vacância financeira e número de cotistas | Preenchimento manual ou importação pelo CSV opcional. |

Os indicadores que exigem um histórico maior do que o disponível ficam em branco. Se a coleta de um fundo falhar, seu ticker continua na planilha e o erro aparece no terminal. Para coletar outros FIIs, edite a lista `FIIS` em `main.py` e informe os códigos sem o sufixo `.SA`.

**Cada execução sobrescreve a planilha anterior.** Se você preencher células diretamente no Excel, guarde uma cópia ou registre esses dados no CSV antes de executar novamente.

## Como ler o grafo

O perfil de cada fundo reúne três indicadores da janela comum de dois anos: rentabilidade acumulada, DY médio e liquidez média diária. A planilha organiza os cálculos em quatro abas:

| Aba | Conteúdo |
| --- | --- |
| `Z_scores` | Os três indicadores padronizados entre fundos com perfil completo. |
| `Correlacao` | A correlação de Pearson entre os perfis de cada par de FIIs. |
| `Pesos_W` | A similaridade do par em escala de 0 a 1, calculada como `(correlação + 1) / 2`. |
| `Pares_do_grafo` | Cada par listado uma vez, com peso, situação e indicação de aresta para pesos a partir de 0,60. |

Para ver as arestas no Excel, abra `Pares_do_grafo` e filtre a coluna **Incluir com limiar atual** pela opção **VERDADEIRO**. A aba **Notas** mostra quantos fundos têm perfil válido, quantos pares têm peso e quantas arestas foram geradas.

O cálculo exige os três indicadores. O DY atual e os indicadores de 5 e 10 anos são descritivos. Pares com dados insuficientes ou correlação indefinida ficam sem peso, com o motivo na coluna **Situação**. Para calcular os DYs mensais, são necessários cerca de três anos de histórico bruto por fundo.

## Como incluir vacância e número de cotistas

O coletor deixa essas colunas em amarelo. Você pode preenchê-las no Excel com dados de relatórios gerenciais, Fundamentus ou Status Invest. Para reaproveitar os dados a cada execução, crie `fundamentos_extras.csv` na mesma pasta de `main.py`:

```csv
Ticker,Vacância física (%),Vacância financeira (%),Cotistas
HGLG11,2.5,1.8,350000
```

Os valores da linha são ilustrativos. Use uma linha por ticker, informe a vacância em pontos percentuais (`2.5` representa 2,5%) e escreva o número de cotistas sem separador de milhar. Execute `python main.py` para incorporar os dados na planilha.
