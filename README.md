# Inventory Helper

Single-page photo-first inventory counter. Take a photo, enter a quantity, name items later.
Counts are kept on the device and, when a Google Sheet is connected, synced to that sheet so
every device sees the same count.

Live page: https://jaysvaletss.github.io/inventory-helper/

## Connecting a Google Sheet

1. Open the sheet, then **Extensions > Apps Script**.
2. Replace the contents of `Code.gs` with [`apps-script/Code.gs`](apps-script/Code.gs).
3. Set `PASSCODE` near the top to a team passcode. Do this only in the Apps Script editor,
   never in this public repo.
4. Pick the `setup` function in the toolbar and press **Run**. Approve the Sheets and Drive
   permission prompts.
5. **Deploy > New deployment > Web app**. Execute as **Me**, who has access **Anyone**. Copy
   the web app URL (it ends in `/exec`).
6. Put that URL in `API_URL` at the top of the script in `index.html` and push.

After changing `Code.gs` later, use **Deploy > Manage deployments > Edit > New version** so
the URL stays the same.

## What ends up where

- **Inventory tab**: one row per item with Name, Qty, Counted at and a Photo link. Editing a
  name or quantity in the sheet shows up on the phones within about 20 seconds. Hidden columns
  to the right hold IDs the page needs; leave them alone.
- **Drive folder "Inventory Helper Photos"**: one JPEG per item, owned by whoever deployed the
  script.
- **Clear the whole count** in the page empties the tab and moves the photos to Drive trash.

Phones keep working with no signal. Changes are saved on the device and sent when the sheet
is reachable again.
