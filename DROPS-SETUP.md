# Plantae drop sales: setup

What changed
- `index.html`: new Plant drop page (`#/drop`), live stock, basket, order form, countdown, home-page strip. Try it without a backend at `#/drop?demo=1`.
- `apps-script/Code.gs`: backend for the Google Sheet. Replaces the sheet's current script (read-only feed) with the same feed plus drop orders.
- `index.before-drops.html`: your previous page, untouched.

Steps (about 10 minutes)
1. Open Plantae_Website_Content_Manager > Extensions > Apps Script. Copy the existing code somewhere safe first.
2. Replace the contents of Code.gs with `apps-script/Code.gs`. Save.
3. Reload the sheet. A "Plantae drops" menu appears. Run "Set up the drop tabs" (approve permissions the first time). It adds Drop Events, Drop Plants, Orders, Order Lines, Dispatch and Drop Settings, and leaves existing tabs alone.
4. File > Settings: set the time zone to Johannesburg, so drop times are read correctly.
5. Deploy > Manage deployments > edit the existing web app > Version: New version > Deploy. The URL stays the same, so the site keeps working.
6. Optional: "Add an example drop" creates a labelled test drop that opens in 2 minutes.
7. Commit `index.html` and push to GitHub Pages.

Running a drop
1. Drop Events: add a row (ID, name, opens, closes).
2. Drop Plants: one row per plant (Drop ID, SKU, quantity; one-offs are 1).
3. Post the photos on WhatsApp with the link to `/#/drop`.
4. Afterwards: Orders tab "Next step" column tells you what to do (add shipping, send invoice, wait for EFT, pack). The Dispatch tab lists paid orders ready to pack.
