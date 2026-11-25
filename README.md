# TradeAnalytics

Offline sales intake starter for a fictional importer. It reads a synthetic
Excel workbook or CSV and summarizes monthly sales units, revenue, and gross
margin by SKU. Inputs use one currency and whole units.

```sh
python3 -m pip install -r requirements.txt
python3 trade_analytics.py examples/sales.xlsx
```

The CSV option has no third-party dependencies:

```sh
python3 trade_analytics.py examples/sales.csv
```

Use the Sales worksheet with date, sku, units, unit_price, and unit_cost headers.
Use ISO dates or Excel date cells. Margin in the starter uses the supplied unit
cost. An Excel-only input can populate sales summaries, but reorder planning
also needs current stock, inbound orders, and lead time.

Proposed work: extend the workbook with Inventory, Purchase Orders, and Import
Costs sheets. Add a local dashboard, deterministic seasonal comparisons and
reorder rules, and landed-cost scenarios with clearly stated user assumptions.
All fixtures are synthetic.

## License

[MIT](LICENSE)
