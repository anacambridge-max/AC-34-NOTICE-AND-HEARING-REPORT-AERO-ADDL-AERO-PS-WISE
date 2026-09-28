# AC-34 Notice & Hearing Dashboard

Vercel-ready Next.js dashboard for AC-34 Matiala SIR-2026.

## Workflow
1. Keep the one-time master mapping: PS -> AERO/Ad.AERO -> BLO Supervisor -> BLO.
2. Upload Previous ECI Excel.
3. Upload Latest ECI Excel.
4. Upload BLO/Other Excel.
5. Dashboard joins all data by PS.
6. Notice Generated uses Latest ECI.
7. Delivered Difference = Latest - Previous.
8. Hearing Difference = Latest - Previous.
9. PS with Latest Hearing Held = 0 stay highlighted red until a later Latest ECI report shows >0.
10. Export the filtered report to Excel.

The app is intentionally client-side, so Excel files are processed in the browser and are not uploaded to a server.
