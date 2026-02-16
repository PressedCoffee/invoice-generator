# "Make a Copy" Link Plan

## How Google Sheets "Make a Copy" Links Work

Google Sheets supports a special URL format that opens a "Make a copy" dialog for any viewer:

```
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/copy
```

When someone clicks this link:
1. Google asks them to sign in (if not already)
2. Shows a "Make a copy" dialog with the sheet name
3. They click "Make a copy" → get their own editable copy in their Drive
4. The copy includes all sheets, formatting, data validation, conditional formatting, AND the Apps Script project

**This is the standard distribution method for Google Sheets templates.**

## Setup Steps (one-time, ~10 minutes)

### 1. Create the Master Template

1. Create a new Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any default code
4. Paste the entire contents of `Code.gs`
5. Save (Ctrl+S)
6. Close the Apps Script editor tab
7. Return to the spreadsheet and refresh the page
8. Click **🧾 Invoice → ⚡ Set Up Template (run this first)**
9. Authorize when prompted (follow the "Advanced → Go to..." flow)
10. Setup completes — all four tabs created with sample data

### 2. Verify Everything Works

1. Go to WorkLog tab, click on an ACME row
2. Click **🧾 Invoice → Generate for Selected Client**
3. Confirm the PDF generates correctly in Drive
4. Delete the test invoice from Drive and clear the Invoices tab row
5. Reset the WorkLog sample rows back to "Unbilled" and clear their InvoiceID values
6. Reset NextInvoiceNumber in Settings back to 1

### 3. Set Sharing Permissions

1. Click **Share** (top right)
2. Under "General access," change to **"Anyone with the link"**
3. Set permission to **"Viewer"** (NOT Editor)
4. Copy the share link — it looks like:
   ```
   https://docs.google.com/spreadsheets/d/ABC123xyz.../edit?usp=sharing
   ```

### 4. Convert to "Make a Copy" Link

Replace `/edit?usp=sharing` with `/copy`:

```
https://docs.google.com/spreadsheets/d/ABC123xyz.../copy
```

This is your distribution link. Put it in the README.

### 5. Update the README

Replace `TEMPLATE_LINK_HERE` in README.md with the `/copy` link.

### 6. Push to GitHub

```bash
cd projects/invoice-generator
git init
git add Code.gs README.md LICENSE
git commit -m "v1.0: Google Sheets Invoice Generator"
git remote add origin https://github.com/YOUR_USERNAME/sheets-invoice-generator.git
git push -u origin main
```

## Important Notes

- The `/copy` link includes the Apps Script. Users get everything in one click.
- Viewers cannot edit your master template. They can only copy it.
- If you update the master template later, existing copies are NOT affected (they're independent).
- To update the published version: edit the master sheet and script, then the `/copy` link automatically serves the latest version.

## Checklist

- [ ] Create Google Sheet
- [ ] Paste Code.gs into Apps Script editor
- [ ] Run setupTemplate()
- [ ] Test invoice generation (then reset sample data)
- [ ] Set sharing to "Anyone with the link → Viewer"
- [ ] Convert share link to /copy link
- [ ] Update README.md with the /copy link
- [ ] Create GitHub repo
- [ ] Push code + README + LICENSE
- [ ] Update README contact email
- [ ] Take 5 screenshots and add to repo
- [ ] Test the /copy link from an incognito window
