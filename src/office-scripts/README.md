# Office Scripts conversion notes

This folder contains an initial Office Script that replicates the Google Apps Script behavior for creating documents (reading the `Documento` sheet and building API payloads) and a helper to apply API results back to the sheet.

Important limitations & recommended approach

- Office Scripts (Excel on the web) cannot reliably make arbitrary external HTTP requests the same way Google Apps Script's `UrlFetchApp` does. To call external APIs securely from Excel:
  - Use Power Automate to call this Office Script. Power Automate can perform HTTP requests (with authentication), and then pass results back to the script to be written to the workbook.
  - Alternatively, host an Azure Function or other service to act as a proxy and call APIs, but Power Automate is the recommended low-code approach.

Workflow pattern (recommended):

1. Run Office Script with action `buildPayloads` (via Power Automate or manually). This returns an array of payloads (one per row to process).
2. In Power Automate, iterate these payloads and call the external API (POST /api/documentos or others) using HTTP connector, collecting results.
3. Call Office Script again with action `applyResults` and pass the results array. The script will update the sheet (NotaCriada and RetornoAPI columns).

This repository contains `DocumentoScript.ts` which implements steps 1 and 3.

Next steps we can perform if you confirm:
- Convert the remaining functions (ValidarPlanilha helpers, criarPedidoDeVenda, etc.) to the same pattern.
- Add a Power Automate Flow example (JSON) that calls the script and the HTTP connector.
- Implement a client-side MD5 if you want the script to build auth payloads (but the actual HTTP call should still run in Power Automate).

If you'd rather have the Office Script directly perform HTTP requests (not recommended), tell me and I will check if your tenant allows outbound requests from Office Scripts and adapt accordingly.
