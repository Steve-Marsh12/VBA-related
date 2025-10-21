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

