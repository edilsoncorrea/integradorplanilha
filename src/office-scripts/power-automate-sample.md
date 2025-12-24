Power Automate sample (high-level)

1) Create an automated flow triggered manually or on schedule.
2) Add action: Run script (Excel Online (Business))
   - Location: OneDrive or SharePoint file where the workbook is stored
   - Document Library / File / Script: select the workbook and choose the script `DocumentoScript` (upload the script to the workbook first via Excel web)
   - Inputs: { "action": "buildPayloads" }
3) The script returns `payloads` array. Use an "Apply to each" to iterate payloads and call HTTP connector to POST to your API endpoint. Authorize with Bearer token (store token in Azure Key Vault or use secure flow).
4) Collect responses into an array of objects: { sheetRow: <sheetRow>, notaCriada: "Sim" | "Não", retorno: "..." }
5) Call the same Office Script again with inputs: { "action": "applyResults", "results": <the results array> }

This pattern keeps API calls in Power Automate (where authentication is easier) and uses Office Scripts purely for spreadsheet IO.
