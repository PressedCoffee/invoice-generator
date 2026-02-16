# Demo Script — Google Sheets Invoice Generator (2 minutes)

Use this script to record a screen walkthrough or create screenshots for the README.

---

## Scene 1: The Setup (0:00–0:30)

**Action:** Open the spreadsheet template.

**Show:**
1. Click through each tab briefly:
   - **Settings** — Business name, address, currency, invoice prefix all filled in
   - **Clients** — Two sample clients (Acme Corp, Globex Inc) with contact details
   - **WorkLog** — Five unbilled rows: 3 for Acme, 2 for Globex. Status column shows "Unbilled" highlighted in yellow
   - **Invoices** — Empty. "This is where your invoice records will appear."

**Narration / Caption:**
> "Four tabs. Settings for your business info, Clients for who you bill, WorkLog for your daily entries, and Invoices where records are kept automatically."

📸 **Screenshot 1:** Menu bar showing the 🧾 Invoice custom menu
📸 **Screenshot 2:** WorkLog tab with five unbilled rows

---

## Scene 2: Generate for One Client (0:30–1:00)

**Action:**
1. Click into the **WorkLog** tab
2. Click on any row belonging to "ACME"
3. Click **🧾 Invoice → Generate for Selected Client**
4. Confirmation dialog appears showing:
   - Invoice #: INV-000001
   - Client: Acme Corporation
   - Total: $1,125.00
5. Click OK

**Narration / Caption:**
> "Select any row for the client you want to invoice. One click. The script finds all their unbilled work, calculates the total, generates the PDF."

📸 **Screenshot 3:** Confirmation dialog with invoice details

---

## Scene 3: See the Result (1:00–1:30)

**Action:**
1. Show the **WorkLog** tab — the 3 Acme rows now show:
   - Status: "Invoiced" (green highlight)
   - InvoiceID: "INV-000001"
2. Click the **Invoices** tab — new row:
   - INV-000001 | ACME | 2026-02-15 | 2026-03-17 | $1,125.00 | [link]
3. Click the PDF link — the invoice opens in a new tab
4. Show the PDF: clean layout with business info, client info, line items table, totals

**Narration / Caption:**
> "Work log updated automatically. Invoice recorded with a direct link to the PDF. The PDF is saved in Drive under /Invoices/2026/02/."

📸 **Screenshot 4:** WorkLog after invoicing (rows marked green)
📸 **Screenshot 5:** The generated PDF invoice

---

## Scene 4: Batch Generate (1:30–1:50)

**Action:**
1. Click **🧾 Invoice → Generate for All Unbilled**
2. Dialog: "Found unbilled work for 1 client: Globex Inc (2 items). Proceed?"
3. Click Yes
4. Result: "1 Invoice Generated: INV-000002 — Globex Industries — $1,053.30"

**Narration / Caption:**
> "Or generate for everyone at once. One click, all clients, all invoices."

---

## Scene 5: Idempotency (1:50–2:00)

**Action:**
1. Click **🧾 Invoice → Generate for All Unbilled** again
2. Dialog: "No unbilled work found for any client."

**Narration / Caption:**
> "Run it again — nothing happens. Only unbilled rows are ever processed. No double-billing."

---

## Screenshot List for README

| # | Filename | Description |
|---|----------|-------------|
| 1 | `menu-bar.png` | Custom menu in the spreadsheet toolbar |
| 2 | `worklog-before.png` | WorkLog with unbilled rows (yellow status) |
| 3 | `confirmation-dialog.png` | Invoice generation confirmation dialog |
| 4 | `worklog-after.png` | WorkLog with invoiced rows (green status) |
| 5 | `pdf-output.png` | The generated PDF invoice |
