# Northwind Home Goods — Performance Summary (Apr 2025 – Sep 2026)

> **Sample project — synthetic data, not client work.** Northwind Home Goods is fictional. Figures tie out to `Northwind_Performance_Report.xlsx` (sheet *Summary KPIs* unless noted).

**Bottom line:** the business grew — Sep 2026 net revenue was $96,212, +39% vs Sep 2025 — but growth is increasingly bought through Meta at worsening efficiency, and too few first-time buyers come back.

## Headline numbers (18 months)

| KPI | Value |
|---|---|
| Net revenue | $1,452,337 |
| Orders | 10,386 |
| AOV | $139.84 |
| Repeat-customer rate (2+ orders) | 23.4% |
| Ad spend (Meta + Google) | $288,607 |
| MER (net revenue ÷ ad spend) | 5.03x |
| Blended CAC | $37.57 |
| Meta ROAS / CAC | 2.36x / $66.29 |
| Google ROAS / CAC | 2.84x / $55.66 |
| Refund rate | 7.3% |
| Gross margin | 60.6% |

## What changed

- **Meta is getting more expensive.** Spend rose from $45,156 (first 6 months) to $70,851 (last 6), while last-click ROAS fell from 2.96x to 2.02x and CAC climbed from $52.57 to $78.29. Google held at 3.07x ROAS / $54.98 CAC over the same last 6 months. Break-even ROAS at our 61% gross margin is 1.65x. *(Channels sheet)*
- **Retention is the weak link.** Only 5.2% of new customers order again the month after their first purchase, and 23.4% ever place a second order. Customers whose first order included Bedding repeat at 30.7% vs 17.8% for everyone else. *(Cohorts sheet)*
- **One product drives a disproportionate share of refunds.** Linen Duvet Cover - Queen (BED-001) has a 19.0% refund rate vs a 7.3% store average — $25,875 refunded over the period, $15,862 more than if it refunded at the store average. It is the #2 product by net revenue. *(Products sheet)*
- **Black Friday discounting is deep.** In Nov 2025 discounts were 9.7% of gross sales (vs 3.3% the rest of the period) and gross margin dipped to 57.6%. *(Monthly sheet)*

## Recommendations

1. **Cap Meta prospecting and move ~20% of its budget to Google Shopping + retention.** Meta's last-6-month ROAS of 2.02x is now close to the 1.65x break-even, and its CAC of $78.29 is 42% above Google's. Set a Meta CAC ceiling of $60 and review weekly; run a 4-week geo holdout to measure Meta's true incrementality before scaling again.
2. **Build a 30/60-day second-purchase program.** With month-1 retention at 5.2%, a post-purchase email/SMS flow (care guide → complementary product → time-limited VIP offer) is the cheapest growth lever. Target: lift the repeat-customer rate from 23.4% to 26.4% within two quarters.
3. **Lead acquisition with Bedding.** Bedding-first customers repeat 1.7x as often. Shift prospecting creative and the welcome offer toward sheet sets and pillows.
4. **Fix or pull Linen Duvet Cover - Queen.** Pull return reasons from Shopify for BED-001; if it's sizing/colour, update the PDP (dimensions, real-light photos). If refund rate doesn't drop below 10% within 60 days, pause paid promotion of the SKU.
5. **Make BFCM shallower.** Test a 15–20% tiered offer (spend-more-save-more) instead of a flat 25% code; every 5 points of discount on November's gross sales is ≈$7,276 of margin.

*Method & caveats: last-click attribution from Shopify order source; CAD converted at monthly finance rates; 37 lines with negative quantities quarantined and 23 orders without an email excluded from customer metrics (see Data Quality Log). Full definitions in the workbook README sheet.*
