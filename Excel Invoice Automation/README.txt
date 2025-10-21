Excel Invoice Automation (VBA)

**Highlights**
- Clean, modular VBA (standard modules + class module)
- Typical real-world feature set: customer / item data import, invoice generation, templated invoice sheet population, PDF export, batch processing, validation and logging
- Fully commented code to demonstrate best practices: error handling, meaningful names, separation of concerns, reusability


Project structure & Key Concepts

1. Excel workbook layout (expected):
   - Sheet "Data_Customers" — columns: CustomerID, Name, Address, Email
   - Sheet "Data_Items" — columns: ItemID, Description, UnitPrice
   - Sheet "Data_Invoices" — columns: InvoiceID, CustomerID, Date, Status
   - Sheet "Invoice_Template" — formatted printable invoice with named ranges used by the code (see README sections below)

2. How this project demonstrates practical skills (for CV bullets):
   - "Designed and implemented an Excel VBA solution to automatically generate, validate, and export invoices to PDF, including batch processing and robust error handling."
   - Shows OOP usage in VBA (Class Invoice), modular design, working with ranges, file I/O, and integration with Excel UI.

---

USAGE (quick)
1. Open the provided workbook (or create a workbook with the sheets listed above).
2. Import the modules and class into the VBA editor (Alt+F11).
3. Configure the named ranges on the Invoice_Template sheet (see section below).
4. Run `Main_ShowMenu` or `Main_GenerateInvoiceFromSelection` from the Macros dialog.

NAMED RANGES EXPECTED ON Invoice_Template
- rngInv_Number          : invoice number cell
- rngInv_Date            : invoice date cell
- rngInv_CustomerName    : customer name cell
- rngInv_CustomerAddress : customer address block start
- rngInv_LineStartRow    : starting row index for the first invoice line (integer)
- rngInv_LineCols        : comma separated column indices or names used for: Description, Qty, UnitPrice, LineTotal
- rngInv_Total           : invoice total cell

LICENSE
MIT — include an appropriate LICENSE file for your repo.