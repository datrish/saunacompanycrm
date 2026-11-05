# saunacompanycrm — Sauna Leads Board

A lead board for The Sauna Company. The Google Sheet **Sauna Lead Tracker** (tab `Leads`) is the database.
This Google Apps Script, attached to that sheet, serves a web page where the team updates each lead:
phoned, replied, quote sent, chased, won, lost, quote £, next follow-up, notes.
Every change is written straight into the sheet, with a history of who changed what and when.

| File | What it is |
|---|---|
| `Code.gs` | Server code: reads and writes the `Leads` tab, plus a one-time `setup()` |
| `Index.html` | The board (daily progress, chase-today list, lead list with quick-update buttons) |
| `appsscript.json` | Apps Script settings (London time zone, web app runs as the person using it) |

## First-time install (about 5 minutes)

1. Open the **Sauna Lead Tracker** sheet → **Extensions → Apps Script**.
2. Paste `Code.gs` into the `Code.gs` file. Add an HTML file named `Index` (**+ → HTML**) and paste `Index.html`.
   Under **Project Settings** tick *Show "appsscript.json"* and paste `appsscript.json`. Save.
3. In the function menu pick **setup** → **Run** → allow access. This adds the columns
   Chases / Last chased / Wix ID / Last activity / History, renames old statuses
   (*Not contacted → New*, *Contacted - no quote → Replied - no quote*), and resets the row colours.
4. **Deploy → New deployment → Web app**. Execute as: **User accessing the web app**.
   Who has access: **Anyone with Google account**. Deploy and copy the web-app URL. That is the board's link.
5. Share the sheet with each team member as **Editor**. Only people the sheet is shared with can open the board.
   The first time, Google shows "Google hasn't verified this app": click **Advanced → Go to … (unsafe)**.
   This is your own script, so it's safe.

## Updating the code
Change the files here, paste them into Apps Script, then **Deploy → Manage deployments → Edit → Version: New version**.
The link stays the same.

## New website leads
A scheduled Claude task reads new Wix form submissions every hour (7am–9pm) and adds them to the top of the
`Leads` tab with status *New*. Phone and info@ leads are added on the board with **+ Add a lead**.
