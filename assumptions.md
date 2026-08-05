# IT Services LBO + DCF Model Assumptions

## Provided inputs

- Acquisition date: `2025-09-30`
- Currency: `GBP millions`, except share counts
- Tax rate: `20.0%`
- CAPEX as a percentage of revenue: `1.0%`
- Working capital as a percentage of revenue: `10.0%`
- Yield on senior debt: `5.0%`
- Deal fees as a percentage of transaction value: `1.5%`
- Cash at Bank: `2.3`

## Company financials

| Metric | Dec 24A | Dec 25B | Dec 26F | Dec 27F | Dec 28F | Dec 29F |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Revenue | 9.990 | 15.700 | 19.000 | 22.700 | 27.300 | 32.700 |
| EBITDA | 1.034 | 2.312 | 2.377 | 2.961 | 3.762 | 4.743 |

## Shareholdings

| Holder | Shares | Ownership |
| --- | ---: | ---: |
| Chairman | 521,805.281 | 33.1% |
| CEO | 553,334.301 | 35.1% |
| Sales Director | 102,469.315 | 6.5% |
| Operations Director | 95,000.02474 | 6.0262% |
| Head of US | 0 | 0.0% |
| External shareholders | 165,000.0427 | 10.4666% |
| Employees | 97,999.02532 | 6.2164% |
| Option holders | 40,843.01054 | 2.5908% |
| Total | 1,576,451 | 100.0% |

## Standard modeling convention used

- EBITDA is treated as EBIT for free-cash-flow purposes because no depreciation or amortization forecast was provided.
- Working capital is modeled as `10.0%` of revenue, with year-over-year change deducted in free cash flow.
- CAPEX is modeled as `1.0%` of revenue.

## Undisclosed items

The prompt did not provide enough information to calculate these without inventing assumptions, so the workbook marks them as `Undisclosed`:

- LBO purchase price / entry enterprise value
- LBO entry EBITDA multiple
- Debt quantum at close
- Equity contribution and sponsor returns
- DCF WACC
- DCF terminal / exit EBITDA multiple
- Public comparable company set and median EBITDA multiple
- Precedent transaction set and lower-quartile EBITDA multiple
