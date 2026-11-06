# Sauna Leads Board – rules for every change

Whenever the board changes (new button, field, status, reason, report or payment rule), update ALL of these in the same piece of work:

1. **Code** – `Index.html` / `Code.gs` in this repo. Test on the mock harness, commit, push.
2. **Live board** – load both files into the Apps Script project "Sauna Leads Board", save, Deploy → Manage deployments → Edit → New version.
3. **Google Sheet "Sauna Lead Tracker"** (1FkCvsCwWN8LDoJQeMFwc7qJB0vkwMusE4uVyHXi_pTY):
   - drop-downs (Status, Lost reason) match the board's lists
   - new columns exist with the right header
   - "How to use" tab matches the board
   - "Payments" tab still builds (refreshPaymentsTab)
4. **Daily Guide Google Doc** (1qWYezbRde47SkuaHGtjCRSr5PSnt8irptZhGUZAD5X0) – buttons, routines and review steps match the board.

Never: contact clients automatically, or send reports to anyone except datrish@gmail.com. Every board action keeps the Confirm step.
