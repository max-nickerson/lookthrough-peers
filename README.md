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
`lookthrough_peers_dados.csv` (`;` separated, `,` decimals): flat table, one row per holding path.

## Speed and reliability

- API responses are cached in one SQLite file (`cache_maisretorno.sqlite`). Monthly portfolios are fetched once; the list of available months is refreshed daily. A rerun from cache takes well under 2 minutes.
- Requests use 3 threads with reused connections. Going faster gets the client temporarily blocked by the site (HTTP 403), so speed comes from the cache, not more threads.
- On 403/429/5xx the script waits and retries (5s, 20s, 60s, 120s). Fund-months that still fail are retried once more at the end, one at a time; anything left is listed as `ATENCAO`, and a rerun later only fetches what is missing.

## Notes

- Months where the manager still keeps the portfolio confidential are skipped.
- If a sub-fund has no open portfolio for a month, its most recent open month before that is used (shown in the `mes_carteira` column).
- Data belongs to Mais Retorno / CVM. Check Mais Retorno's terms of use before redistributing anything produced with this script.
