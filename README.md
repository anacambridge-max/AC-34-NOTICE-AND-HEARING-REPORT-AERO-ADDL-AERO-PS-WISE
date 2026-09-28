# AC-34 Notice & Hearing Dashboard

Vercel-ready Next.js dashboard for AC-34 Matiala SIR-2026.

## Daily workflow
1. Upload the one-time Master Mapping Excel containing PS No, AERO/Ad.AERO, AERO mobile, BLO Supervisor, BLO Name and BLO Mobile. The browser saves this mapping in localStorage, so it does not need to be uploaded every day.
2. Upload Previous ECI Excel — comparison baseline.
3. Upload Latest ECI Excel — current values.
4. Upload BLO / Other Excel — documents/anomaly/discrepancy/BLO-letter values.
5. All daily data is joined by PS number.
6. Notice Generated comes from Latest ECI.
7. Delivered Difference = Latest Delivered − Previous Delivered.
8. Hearing Difference = Latest Hearing Held − Previous Hearing Held.
9. Every PS with Latest Hearing Held = 0 is highlighted red. The highlight disappears automatically when a later Latest ECI upload shows Hearing Held > 0.
10. AERO/Ad.AERO and BLO Supervisor summaries update automatically.
11. Filter by AERO, Supervisor, PS/BLO search and export the current view to Excel.

## Vercel deployment
- Import this GitHub repository into Vercel.
- Framework: Next.js (auto-detected).
- Build command: next build.
- No environment variables are required.
- Deploy.

The Excel files are processed in the browser; this app does not send the uploaded Excel files to a server.

## Important
The master mapping is intentionally separated from daily ECI metrics. This prevents a new ECI report from changing the PS-to-officer/BLO relationship.
