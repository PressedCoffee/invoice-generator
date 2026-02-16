# 🧾 Google Sheets Invoice Generator

**Free, open-source invoice generation that lives inside your spreadsheet.**

Turn your Google Sheets work log into professional PDF invoices with one click. No SaaS subscription, no third-party accounts, no data leaving your Google account.

![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)

---

## What It Does

- Log your work in a simple spreadsheet (date, client, description, qty, rate)
- Click **Invoice → Generate** from the menu
- Get a formatted PDF invoice saved to your Google Drive
- Work log rows are automatically marked as invoiced — no double-billing

## Features

- ✅ **One-click setup** — run `setupTemplate()` and all sheets are created with formatting, validation, and sample data
- ✅ **One-click PDF generation** from a custom menu
- ✅ **Generate per client** or **batch-generate for all unbilled work**
- ✅ **Auto-incrementing invoice numbers** (configurable prefix)
- ✅ **Organized Drive storage**: `/Invoices/YYYY/MM/INV-000123_ClientName_2026-02-15.pdf`
- ✅ **Tax support** per line item (percentage-based)
- ✅ **Multi-currency** (USD, EUR, GBP, CAD, AUD, and more)
- ✅ **Idempotent** — running twice won't double-bill; only unbilled rows are processed
- ✅ **No external dependencies** — runs entirely within Google Workspace
- ✅ **Free forever** — MIT licensed

---

## Setup (3 steps, ~5 minutes)

### Step 1: Copy the Spreadsheet

**Option A — Use the template (recommended):**

👉 **[Click here to make a copy](TEMPLATE_LINK_HERE)** — opens a ready-to-use spreadsheet in your Google Drive.

**Option B — Start from scratch:**

1. Create a new Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any default code in the editor
4. Copy the entire contents of [`Code.gs`](Code.gs) and paste it in
5. Click **💾 Save** (or Ctrl+S)
6. Close the Apps Script editor
7. Refresh the spreadsheet
8. Click **🧾 Invoice → ⚡ Set Up Template (run this first)**

The setup function creates all four sheets (Settings, Clients, WorkLog, Invoices) with proper headers, formatting, data validation, conditional formatting, and sample data. It's safe to run multiple times — it skips sheets that already exist.

### Step 2: Authorize

The first time you run any function, Google will ask for permissions:

1. Click **Continue** on the authorization prompt
2. Choose your Google account
3. You'll see "Google hasn't verified this app" — click **Advanced → Go to Invoice Generator (unsafe)**
4. Click **Allow**

The script needs permission to:
- Read and edit your spreadsheet (to process work log entries)
- Create files in Google Drive (to save invoice PDFs)
- Create Google Docs (temporary, used to generate PDFs — auto-deleted)

> **Your data never leaves your Google account.** There are no external servers, analytics, or tracking.

### Step 3: Configure & Generate

1. Open the **Settings** tab — replace the placeholder values with your business info:

| Setting | What to enter |
|---------|--------------|
| BusinessName | Your business or personal name |
| BusinessAddress | Your address (use Shift+Enter for multiple lines) |
| BusinessEmail | Your billing email |
| Currency | USD, EUR, GBP, etc. |
| InvoicePrefix | e.g., `INV-` or `2026-` |
| NextInvoiceNumber | Starting number (e.g., `1`) |
| DefaultDueDays | Days until payment due (e.g., `30`) |
| InvoiceNotes | Optional footer text on invoices |

2. Open the **Clients** tab — replace sample clients with yours. The `ClientID` is a short code you'll use in the work log.

3. Open the **WorkLog** tab — delete the sample rows and start logging your work. Set Status to `Unbilled`.

4. When ready, click **🧾 Invoice → Generate for Selected Client** (click on a row first) or **Generate for All Unbilled**.

Your PDF appears in Google Drive under `/Invoices/YYYY/MM/`.

---

## Sheet Structure

| Tab | Purpose | You edit? |
|-----|---------|-----------|
| **Settings** | Business info and invoice configuration | ✅ Yes — fill in your details |
| **Clients** | Client directory (ID, name, email, address, terms) | ✅ Yes — add your clients |
| **WorkLog** | Daily work entries with billing status | ✅ Yes — log work here |
| **Invoices** | Auto-generated record of all invoices with PDF links | ❌ No — populated by the script |

---

## How It Works

```
You log work in WorkLog (Status = "Unbilled")
        ↓
Click 🧾 Invoice → Generate
        ↓
Script reads unbilled rows for the client
        ↓
Calculates line items, tax, totals
        ↓
Creates a formatted Google Doc → exports as PDF
        ↓
Saves PDF to Drive: /Invoices/YYYY/MM/INV-000123_Client_Date.pdf
        ↓
Marks WorkLog rows as "Invoiced" + stores InvoiceID
        ↓
Records the invoice in the Invoices tab with PDF link
        ↓
Deletes the temporary Google Doc
```

**Idempotency:** Only rows with Status = `Unbilled` are ever processed. If you click generate twice, the second run finds nothing to bill. No duplicates, no double-charging.

---

## Demo (2 minutes)

> **0:00–0:30 — The Setup:** Open the spreadsheet. Four tabs: Settings (filled in), Clients (two sample clients), WorkLog (five unbilled rows), Invoices (empty).
>
> **0:30–1:00 — Generate One:** Click a row for "ACME." Click 🧾 Invoice → Generate for Selected Client. Dialog shows INV-000001, Acme Corporation, $1,125.00.
>
> **1:00–1:30 — See the Result:** WorkLog rows turn green with invoice ID. Invoices tab has a new row with PDF link. Click it — clean professional invoice.
>
> **1:30–1:50 — Batch:** Click Generate for All Unbilled. Globex's 2 items get invoiced in one click.
>
> **1:50–2:00 — Idempotent:** Click Generate for All Unbilled again. "No unbilled work found." Nothing happens.

---

## Screenshots

### 1. The Custom Menu
*The 🧾 Invoice menu with setup and generate options.*

`[screenshot: menu-bar.png]`

### 2. WorkLog Before Invoicing
*Five unbilled rows (yellow) across two clients.*

`[screenshot: worklog-before.png]`

### 3. Generation Confirmation
*Invoice number, client, and total displayed before PDF is created.*

`[screenshot: confirmation-dialog.png]`

### 4. WorkLog After Invoicing
*Rows marked as "Invoiced" (green) with invoice IDs filled in.*

`[screenshot: worklog-after.png]`

### 5. Generated PDF Invoice
*Clean, professional PDF with line items, tax, and totals.*

`[screenshot: pdf-output.png]`

---

## File Organization

Invoices are saved to Google Drive automatically:

```
My Drive/
  Invoices/
    2026/
      01/
        INV-000001_Acme_Corporation_2026-01-15.pdf
        INV-000002_Globex_Industries_2026-01-28.pdf
      02/
        INV-000003_Acme_Corporation_2026-02-15.pdf
```

---

## FAQ

**Can I customize the invoice layout?**
The current version produces a clean, professional document. Custom themes are planned for a future release.

**What if I run it twice?**
Nothing bad happens. Only `Unbilled` rows are processed. Already-invoiced rows are skipped.

**Can it email the invoice to clients?**
Not in v1.0. Email sending is planned for v1.1.

**Does it support multiple currencies?**
Yes. Set your currency code in Settings. Supports USD, EUR, GBP, CAD, AUD, JPY, CNY, INR, BRL, MXN, and any other code.

**Is my data sent anywhere?**
No. Everything runs inside your Google account. No external servers, no analytics, no tracking.

**Can I use this for my business?**
Yes — MIT licensed, free for any use. But this is not accounting software. Consult a tax professional for compliance with your local invoicing requirements.

---

## Roadmap

- [x] v1.0 — Core invoice generation with PDF output + one-click setup
- [ ] v1.1 — Email sending (send invoice directly to client)
- [ ] v1.2 — Payment status tracking (Paid/Unpaid/Overdue)
- [ ] v1.3 — Recurring invoice templates
- [ ] v2.0 — Dashboard with revenue charts

---

## Contributing

Found a bug? Have a feature idea? [Open an issue](../../issues) or submit a PR.

---

## License

MIT License. See [LICENSE](LICENSE) for details.

**Disclaimer:** This tool generates invoices for convenience. It is not accounting software. Consult a tax professional for compliance with your local invoicing requirements. The authors are not responsible for errors in generated invoices.

---

**Need this customized for your business?** [Get in touch](mailto:CONTACT_EMAIL_HERE) — I build automation tools for freelancers and small businesses.
