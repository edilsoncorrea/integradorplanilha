/**
 * Versão do DocumentoScript adaptada para testes
 * Exporta funções individuais para facilitar testes unitários
 */

// Constantes (mantidas iguais)
const HOST = "https://087344bimerapi.alterdata.cloud";
const CodigoDaEmpresa = 0;
const CodigoCliente = 1;
const NomeDoCliente = 2;
const IdentificadorCliente = 3; 
const CodigoDaOperacao = 4; 
const IdentificadorOperacao = 5; 
const CFOP = 6; 
const CodigoDoServico = 7; 
const IdentificadorServico = 8; 
const NomeDoServico = 9;
const Quantidade = 10;
const Valor = 11; 
const Descriminacao1 = 12; 
const Descriminacao2 = 13; 
const Codigoprazo = 14;
const IdentificadorPrazo = 15;
const FormaPagamentoEntrada = 16;
const CodidoDaFormaDePagamento = 17;
const IdentificadorFormaPagamento = 18;
const DataEmissao = 19;
const VencimentoFatura = 20; 
const NotaCriada = 21;
const RetornoAPI = 22; 

export function main(workbook: any, inputs?: any): any {
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

export function buildPayloads(workbook: any) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };

  const used = sheet.getUsedRange();
  if (!used) return { payloads: [] };
  const values = used.getValues() as any[][];

  const payloads: any[] = [];

  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const notaCriada = row[NotaCriada];
    if (!notaCriada || String(notaCriada).trim() === '') {
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

      payloads.push({ sheetRow: i + 1, payload });
    }
  }

  return { payloads };
}

export function buildObservacao(row: any[]) {
  const ValorDocumento = formatCurrency(row[Valor]);
  const observacao = `${row[NomeDoServico]}(${row[Quantidade]} X ${ValorDocumento}) - ${row[Valor]}\n\n\n` +
    `${row[Descriminacao1]}\n\n${row[Descriminacao2]}\n\n Data Vencimento: ${row[VencimentoFatura]}`;
  return observacao;
}

export function formatCurrency(value: any) {
  if (typeof value === 'number') {
    return value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  }
  return value;
}

export function applyResults(workbook: any, results: Array<{sheetRow:number, notaCriada?: string, retorno?: string}>) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };

  for (const res of results) {
    const row = res.sheetRow;
    if (!row || row < 1) continue;
    const nota = res.notaCriada || '';
    const retorno = res.retorno || '';

    sheet.getRangeByIndexes(row - 1, NotaCriada, 1, 1).setValue(nota);
    sheet.getRangeByIndexes(row - 1, RetornoAPI, 1, 1).setValue(retorno);
  }

  return { ok: true, updated: results.length };
}