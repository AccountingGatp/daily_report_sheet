# Saint Daily Reporting

Static site — no build step, no server code, no login.

## Structure

    index.html      the app (root, so Vercel serves it at /)
    js/
      sales-engine.js    sales reconciliation engine (Shopify Net payments + Total sales)
      daily-engine.js    daily report calculations, file parsing and checks
      report-export.js   builds the Excel workbook (ExcelJS)
    vercel.json     static config: never cache js/ so updates go live immediately

SheetJS and ExcelJS load from cdn.jsdelivr.net, so the browser needs internet access.

## Deploy to Vercel

Option A — dashboard:
1. Push this folder to a Git repo (GitHub/GitLab/Bitbucket).
2. Vercel → Add New → Project → import the repo.
3. Framework Preset: Other. Build Command: leave empty. Output Directory: leave empty
   (the repo root already holds index.html). Root Directory: this folder, if the repo has more in it.
4. Deploy.

Option B — CLI:
    npm i -g vercel
    cd saint-daily
    vercel          # preview
    vercel --prod   # production

## Settings

Links to the other Saint tools are at the top of index.html:

    window.TOOL_LINKS = {
      cogs:     'https://cogs-app.vercel.app/',
      merchant: 'https://merchant-fee-app.vercel.app/',
      sales:    'https://saintsalesautomation.vercel.app/'
    };

Change a link there and redeploy; the SOP panel picks it up.

## Notes

- Everything runs in the browser. No file is uploaded to a server.
- To run it locally, serve the folder (e.g. `python3 -m http.server`) and open the address.
  Opening index.html directly from disk also works in Chrome and Edge.
