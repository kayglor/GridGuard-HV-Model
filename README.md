# GridGuard HV Model

Concept-stage financial model and enterprise plan for **GridGuard HV Services**, a service business that sells preventive and corrective maintenance for 11 kV boards, underground cables and 11/0.415 kV transformers.

> **All figures are assumptions, not forecasts.** Prices, hours, salaries, overheads and site counts are illustrative. Currency is generic currency units (CU). No company has been registered, and nothing here is legal, tax, financial or engineering advice. Test limits and standards references are indicative: confirm them with the manufacturer, the applicable standards and local regulations.

## What is in this repository

| Path | Contents |
|---|---|
| `model/GridGuard_HV_Enterprise_Model.xlsx` | Three-year model: Inputs, Projections (revenue, costs, profit, staff capacity, cash) and Summary. Formula-driven; edit the blue cells on the Inputs sheet. |
| `model/GridGuard_HV_Enterprise_Plan.rtf` | Enterprise plan: services, customers, operating model, team, compliance checklist, financial summary, go-to-market, risks and a 90-day launch plan. |
| `src/enterprise.js` | The assumptions (sites, services, prices, team, overheads) and the calculation used for cached results. |
| `src/build_enterprise.js` | Builds the Excel model and the RTF plan from `src/enterprise.js`. |

## Rebuild the files

Requires Node.js 18 or later.

```bash
npm install
npm run build
```

This rewrites the two files in `model/`.

## Reference site

Each customer site is modelled like one reference estate: 10 transformers (500 kVA each), 10 cable circuits and 1 11 kV board.

## Headline assumed results

| | Year 1 | Year 2 | Year 3 |
|---|---|---|---|
| Sites under contract | 5 | 14 | 22 |
| Revenue (CU) | 428,000 | 1,157,102 | 1,775,098 |
| Net profit (CU) | -121,440 | 77,233 | 164,995 |
| Cumulative cash (CU) | -265,440 | -200,207 | -51,213 |

Year 1 is an investment year with a peak funding need of about 265,000 CU. Tax has no loss carry-forward, and working capital and financing costs are not modelled.

## Safety note

AI and automation in this concept are advisory only. Isolation, switching, return to service, permits and safety sign-off stay with a competent person.

## Status

Concept. No license has been chosen for this repository, so by default all rights are reserved. Add a `LICENSE` file if you want others to reuse it.
