# lookthrough-peers

Look-through of Brazilian fund portfolios using Mais Retorno's public portfolio data (sourced from CVM/CDA).

For each fund in `PEERS`, every monthly portfolio from the oldest to the latest open month is fetched. Every fund-quota holding whose sub-fund publishes a portfolio is opened recursively, and percentages multiply down the chain (90% sub-fund x 20% asset = 18% of the top fund). Sub-funds without a portfolio stay as a single line in their category (FIDC, FIAGRO or Cotas de Fundos).

## Usage

Edit `PEERS` (fund name -> Mais Retorno link or CNPJ) at the top of the script, then:

```bash
pip install pandas requests openpyxl
python lookthrough_peers.py
```

## Output: `lookthrough_peers.xlsx`

- `graficos`: stacked-area chart of % of NAV by category over time, one per fund
- `evolucao`: the category-by-month table behind the charts
- `top10`: top 10 assets per category across all funds (latest month of each)
- one sheet per fund: month > category > assets, as collapsible row groups
- `dados`: flat table, one row per holding path

API responses are cached in `cache_maisretorno/`, so reruns only fetch new months.

## Notes

- Months where the manager still keeps the portfolio confidential are skipped.
- If a sub-fund has no open portfolio for a month, its most recent open month before that is used (shown in the `mes_carteira` column).
- Data belongs to Mais Retorno / CVM. Check Mais Retorno's terms of use before redistributing anything produced with this script.
