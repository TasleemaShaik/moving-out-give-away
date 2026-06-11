Static giveaway page with shared claim status backed by Google Sheets.

Setup:

1. Open [google-apps-script/Code.gs](google-apps-script/Code.gs) and copy it into a Google Apps Script project that is bound to your Google Sheet.
2. In that sheet, create or keep a tab named `Claims`.
3. Run `syncItemsToSheet()` once from Apps Script to add all 37 items with `Available` status.
4. Deploy the Apps Script as a web app.
   Execute as: `Me`
   Who has access: `Anyone`
5. Copy the deployed web app URL.
6. Open [index.html](index.html) and set `const CLAIMS_API_URL = "YOUR_WEB_APP_URL"` if it is not already configured.
7. Push this folder to GitHub Pages.

Sheet columns used by the script:

`item_id | item_name | status | claimed_by | claimed_at`

Notes:

- `index.html` is the live GitHub Pages version with shared claim status.
- Item IDs are based on the image filenames in `images/`.
