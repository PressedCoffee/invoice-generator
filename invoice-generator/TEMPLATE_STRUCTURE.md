# Template Sheet Structure
# 
# This file documents the exact sheet structure for the Google Sheets template.
# Use this to manually create the template or to verify the template is correct.
#
# The template should be created as a Google Sheet and shared with 
# "Anyone with the link can view" + "Make a copy" instructions.

## =============================================
## TAB 1: Settings
## =============================================
## 
## Column A: Key (String)
## Column B: Value (String/Number)
## 
## Row 1 (Header): Key | Value
## Row 2: BusinessName | [Your Business Name]
## Row 3: BusinessAddress | [123 Main St\nCity, ST 00000]
## Row 4: BusinessEmail | [you@example.com]
## Row 5: LogoURL | [optional - for future use]
## Row 6: Currency | USD
## Row 7: InvoicePrefix | INV-
## Row 8: NextInvoiceNumber | 1
## Row 9: DefaultDueDays | 30
## Row 10: InvoiceNotes | [optional - appears at bottom of invoice]
##
## Formatting:
## - Row 1: Bold, light gray background (#F3F3F3)
## - Column A: Width 200px
## - Column B: Width 400px

## =============================================
## TAB 2: Clients
## =============================================
## 
## Columns: ClientID | Name | Email | Address | PaymentTerms
## 
## Row 1 (Header): ClientID | Name | Email | Address | PaymentTerms
## Row 2 (Sample): ACME | Acme Corporation | billing@acme.com | 456 Oak Avenue\nSan Francisco, CA 94102 | Net 30
## Row 3 (Sample): GLOBEX | Globex Industries | ap@globex.com | 789 Pine Street\nLos Angeles, CA 90001 | Net 15
##
## Formatting:
## - Row 1: Bold, light gray background (#F3F3F3), frozen
## - ClientID column: Width 100px
## - Name column: Width 200px
## - Email column: Width 200px
## - Address column: Width 250px
## - PaymentTerms column: Width 150px

## =============================================
## TAB 3: WorkLog
## =============================================
## 
## Columns: Date | ClientID | Description | Qty | Rate | Tax % | Status | InvoiceID
## 
## Row 1 (Header): Date | ClientID | Description | Qty | Rate | Tax % | Status | InvoiceID
## Row 2 (Sample): 2026-02-01 | ACME | Website redesign - homepage | 8 | 75.00 | 0 | Unbilled |
## Row 3 (Sample): 2026-02-03 | ACME | Website redesign - about page | 4 | 75.00 | 0 | Unbilled |
## Row 4 (Sample): 2026-02-05 | ACME | Website redesign - contact form | 3 | 75.00 | 0 | Unbilled |
## Row 5 (Sample): 2026-02-02 | GLOBEX | Logo design - 3 concepts | 1 | 500.00 | 8.5 | Unbilled |
## Row 6 (Sample): 2026-02-04 | GLOBEX | Brand guidelines document | 6 | 85.00 | 8.5 | Unbilled |
##
## Formatting:
## - Row 1: Bold, light gray background (#F3F3F3), frozen
## - Date column: Date format (yyyy-MM-dd), width 110px
## - ClientID: Width 100px
## - Description: Width 300px
## - Qty: Number format (0.##), width 60px, right-aligned
## - Rate: Currency format, width 80px, right-aligned
## - Tax %: Percentage format (0.#%), width 70px, right-aligned
## - Status: Width 90px; Data validation dropdown: Unbilled, Invoiced
## - InvoiceID: Width 120px
##
## Conditional formatting:
## - Status = "Unbilled": light yellow background (#FFF9C4)
## - Status = "Invoiced": light green background (#C8E6C9)

## =============================================
## TAB 4: Invoices
## =============================================
## 
## Columns: InvoiceID | ClientID | IssueDate | DueDate | Total | PDF_Link
## 
## Row 1 (Header): InvoiceID | ClientID | IssueDate | DueDate | Total | PDF_Link
## (No sample data — this is populated by the script)
##
## Formatting:
## - Row 1: Bold, light gray background (#F3F3F3), frozen
## - InvoiceID: Width 120px
## - ClientID: Width 100px
## - IssueDate: Date format, width 110px
## - DueDate: Date format, width 110px
## - Total: Currency format, width 100px, right-aligned
## - PDF_Link: Width 300px
