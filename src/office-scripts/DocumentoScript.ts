/**
 * DocumentoScript.ts
 *
 * Office Script version of the Google Apps Script document-creation flow.
 * Pattern used:
 *  - action = 'buildPayloads' -> returns array of payloads to call API
 *  - action = 'applyResults' -> receives an array of {row: number, notaCriada: string, retorno: string} and writes them to the sheet
 *
 * Usage (recommended): call this script from Power Automate.
 */

/**
 * Index constants (copied from Constantes.gs). Columns are 0-based when reading getValues().
 */
const HOST = "https://087344bimerapi.alterdata.cloud";
const CodigoDaEmpresa = 0;
const CodigoCliente	= 1;
const NomeDoCliente	= 2;
const IdentificadorCliente = 3; 
const CodigoDaOperacao = 4; 
const IdentificadorOperacao = 5; 
const CFOP = 6; 
const CodigoDoServico = 7; 
const IdentificadorServico = 8; 
const NomeDoServico	= 9;
const Quantidade= 10;
const Valor = 11; 
const Descriminacao1 = 12; 
const Descriminacao2 = 13; 
const Codigoprazo = 14;
const IdentificadorPrazo = 15
const FormaPagamentoEntrada = 16
const CodidoDaFormaDePagamento = 17;
const IdentificadorFormaPagamento	= 18;
const DataEmissao	= 19;
const VencimentoFatura = 20; 
const NotaCriada = 21;
const RetornoAPI	= 22; 

// Office Script entry point. 'inputs' is optional and is used when calling from Power Automate.
export function main(workbook: ExcelScript.Workbook, inputs?: any): any {
  const action = inputs && inputs.action ? inputs.action : 'buildPayloads';

  if (action === 'buildPayloads') {
    return buildPayloads(workbook);
  } else if (action === 'applyResults') {
    const results = inputs && inputs.results ? inputs.results : [];
    return applyResults(workbook, results);
  } else {
    return { error: 'unknown action ' + action };
  }
}

function buildPayloads(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };

  const used = sheet.getUsedRange();
  if (!used) return { payloads: [] };
  const values = used.getValues() as any[][];

  // We'll return payloads with the row number (1-based) so Power Automate can correlate responses
  const payloads: any[] = [];

  for (let i = 2; i < values.length; i++) { // start at row index 2 to skip header rows (matches GAS)
    const row = values[i];
    const notaCriada = row[NotaCriada];
    if (!notaCriada || String(notaCriada).trim() === '') {
      // Build payload similar to criarDocumento / criarPedidoDevenda3
      const payload: any = {
        StatusNotaFiscalEletronica: 'A',
        TipoDocumento: 'S',
        TipoPagamento: '0',
        CodigoEmpresa: row[CodigoDaEmpresa],
        DataEmissao: row[DataEmissao],
        DataReferencia: row[DataEmissao],
        DataReferenciaPagamento: row[DataEmissao],
        IdentificadorOperacao: row[IdentificadorOperacao],
        IdentificadorPessoa: row[IdentificadorCliente],
        Observacao: buildObservacao(row),
        Itens: [
          {
            CFOP: row[CFOP],
            IdentificadorProduto: row[IdentificadorServico],
            Quantidade: row[Quantidade],
            ValorUnitario: row[Valor]
          }
        ],
        Pagamentos: [
          {
            Aliquota: 100,
            AliquotaConvenio: 10,
            DataVencimento: row[DataEmissao],
            IdentificadorFormaPagamento: row[IdentificadorFormaPagamento],
            Valor: row[Valor]
          }
        ]
      };

      payloads.push({ sheetRow: i + 1, payload }); // sheetRow is 1-based row number
    }
  }

  return { payloads };
}

function buildObservacao(row: any[]) {
  const ValorDocumento = formatCurrency(row[Valor]);
  const observacao = `${row[NomeDoServico]}(${row[Quantidade]} X ${ValorDocumento}) - ${row[Valor]}\n\n\n` +
    `${row[Descriminacao1]}\n\n${row[Descriminacao2]}\n\n Data Vencimento: ${row[VencimentoFatura]}`;
  return observacao;
}

function formatCurrency(value: any) {
  // Basic formatting for PT-BR currency; Power Automate / API may expect numbers, so keep numeric value too
  if (typeof value === 'number') {
    return value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  }
  return value;
}

function applyResults(workbook: ExcelScript.Workbook, results: Array<{sheetRow:number, notaCriada?: string, retorno?: string}>) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };

  // For each result, write NotaCriada and RetornoAPI into the appropriate columns
  for (const res of results) {
    const row = res.sheetRow; // 1-based
    if (!row || row < 1) continue;
    const nota = res.notaCriada || '';
    const retorno = res.retorno || '';

    // NotaCriada is column index NotaCriada (0-based) -> cell column is NotaCriada + 1
    sheet.getRangeByIndexes(row - 1, NotaCriada, 1, 1).setValue(nota);
    sheet.getRangeByIndexes(row - 1, RetornoAPI, 1, 1).setValue(retorno);
  }

  return { ok: true, updated: results.length };
}
