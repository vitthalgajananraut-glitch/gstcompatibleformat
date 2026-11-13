# GST Sales Register Converter

A browser-based tool for converting ecommerce sales reports into a generic Indian GST sales register CSV. CSV and Excel (`.xls` / `.xlsx`) files are read locally in the browser; report data is not sent to a server.

## Run

Open `index.html` in a modern browser. The spreadsheet reader (SheetJS) is loaded from jsDelivr, so an internet connection is needed to load the page's XLS/XLSX support.

## Use

1. Select the CSV or Excel sales report. For workbooks with multiple sheets, select the worksheet containing the report.
2. Select the seller's state or union territory.
3. Check and correct the automatically suggested source-column mappings.
4. Choose **Convert report**, review the preview and any row-level errors, then download the CSV.

The output includes invoice details, customer GSTIN/name/state, place of supply, HSN/SAC, taxable value, GST rate, CGST, SGST, IGST, cess, and invoice total. GST is calculated only when source tax amounts are absent and both a GST rate and place of supply are available. For same-state supplies, GST is split equally between CGST and SGST, with rounding to two decimal places; for other-state supplies it is assigned to IGST. Existing source tax amounts and invoice totals are preserved when mapped.

Rows with missing required values, invalid numbers, or GSTIN format/known-state-code issues are reported and excluded from the downloadable output. Missing place of supply is inferred from the mapped customer state or the state code in the customer GSTIN. Review all mappings, calculations, and output with a qualified tax professional; the generated register is not a GST portal filing or tax-compliance guarantee.
