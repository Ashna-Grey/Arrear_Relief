# Arrear Relief Calculator – Section 157 (Form 39)

A simple static website (HTML + CSS + JavaScript, no backend) to estimate arrear relief
under Section 157 of the Income-tax Act, 2025. All calculation runs in the browser.

## Files
- `index.html` – page layout and help cards
- `style.css` – styling (light/dark, mobile friendly)
- `script.js` – calculation logic (matches the Excel formulas)

## Deploy on GitHub Pages
1. Create a new repository on GitHub (e.g. `arrear-relief-calculator`).
2. Upload all files from this folder to the repository root (drag & drop in the web UI works).
3. Go to **Settings → Pages**. Under *Build and deployment*, choose **Deploy from a branch**,
   select branch `main` and folder `/ (root)`, then **Save**.
4. After a minute your site is live at `https://<your-username>.github.io/<repo-name>/`.

## Formulas used
- Taxable income after arrears = income before + arrears
- Extra tax in receipt year = max(0, tax after − tax before)
- Extra tax for each earlier year = max(0, tax with − tax without)
- Relief = max(0, receipt-year extra tax − total earlier-year extra tax)
- Net extra tax = max(0, receipt-year extra tax − relief)

Indicative only. Confirm with your DDO or tax adviser before filing Form 39.
