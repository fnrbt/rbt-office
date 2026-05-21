# Real-world Excel parity targets

Fancy, formatting-heavy real-world spreadsheets used as end-to-end render-parity
goals (make our render match Excel, fixed incrementally and logged in
docs/excel-parity-fixes.html). NOT part of the smoke/parity gate corpus.

Real-estate investment / financial models (formatting-rich, NO charts — fully
renderable, ideal for incremental cell-formatting fixes):

| file | source | sheets | notable features |
|------|--------|--------|------------------|
| lyric-multifamily.xlsx | github.com/mdcnuydt/financial-modelling lyric_multifamily/model.xlsx | 4 | colored section bands, 33 fonts, 19 num-fmts, 28 borders, 10 fills, merges, data validation |
| jumeirah-hotel.xlsx | mdcnuydt/financial-modelling jumeirah_beach_hotel/model.xlsx | 1 | compact pro forma, 30 fonts, 26 borders |
| milk-street-office.xlsx | mdcnuydt/financial-modelling 45_milk_street_office/model.xlsx | 2 | office acquisition model |
| saopaulo-condos.xlsx | mdcnuydt/financial-modelling sao_paulo_condos/model.xlsx | 2 | includes a conditional-formatting sheet |
| conditional-formatting-tutorial.xlsx | github.com/AlexTheAnalyst/Excel-Tutorial | 2 | focused conditional-formatting cases |

Stretch — colorful but chart-dominated (saved for later; blocked on chart
rendering, which isn't first-class yet):

| file | source | features |
|------|--------|----------|
| dashboard-ecommerce-sales.xlsx | github.com/chahat-7/Sales-Analysis-Dashboard-Excel | 10 sheets, 69 charts, 15 pivots, 11 images |
| dashboard-sales-analysis.xlsx | github.com/AD-DSandML/sales_analysis_dashboard_in_excel | 8 sheets, 18 charts, 8 pivots, 1 sparkline + conditional formatting |
| dcf-stock-analysis.xlsx | github.com/LondonMarket/Financial-Model-Excel-Template | 2 sheets, DCF/stock template, 24 merged cells |
