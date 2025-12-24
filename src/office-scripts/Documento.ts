// Documento.ts
// Conversão de Documento.gs para Office Scripts (TypeScript)
// Actions:
//  - action: 'buildDocumentoPayloads' -> retorna lista de payloads { sheetRow, payload }
//  - action: 'applyDocumentoResults' -> recebe results array e escreve NotaCriada / RetornoAPI

import {
  CodigoDaEmpresa, IdentificadorOperacao, IdentificadorCliente, NomeDoServico,
  Quantidade, Valor, Descriminacao1, Descriminacao2, VencimentoFatura, DataEmissao,
  CFOP, IdentificadorServico, IdentificadorFormaPagamento, NotaCriada, RetornoAPI
} from './Constantes';

export function main(workbook: any, inputs?: any): any {
  const action = (inputs && inputs.action) ? inputs.action : 'buildDocumentoPayloads';

  if (action === 'buildDocumentoPayloads') {
    const sheet = workbook.getWorksheet('Documento');
    if (!sheet) return { error: 'Worksheet Documento not found' };
    const used = sheet.getUsedRange();
    if (!used) return { payloads: [] };
    const values = used.getValues() as any[][];

    const payloads: any[] = [];
    for (let i = 2; i < values.length; i++) {
      const row = values[i];
      if (!row[NotaCriada] || String(row[NotaCriada]).trim() === '') {
        const observacao = buildObservacao(row);

        const payload = {
          StatusNotaFiscalEletronica: 'A',
          TipoDocumento: 'S',
          TipoPagamento: '0',
          CodigoEmpresa: row[CodigoDaEmpresa],
          DataEmissao: row[DataEmissao],
          DataReferencia: row[DataEmissao],
          DataReferenciaPagamento: row[DataEmissao],
          IdentificadorOperacao: row[IdentificadorOperacao],
          IdentificadorPessoa: row[IdentificadorCliente],
          Observacao: observacao,
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

  if (action === 'applyDocumentoResults') {
    const results = (inputs && inputs.results) ? inputs.results : [];
    const sheet = workbook.getWorksheet('Documento');
    if (!sheet) return { error: 'Worksheet Documento not found' };

    for (const res of results) {
      const row = res.sheetRow;
      if (!row) continue;
      const nota = res.notaCriada || '';
      const retorno = res.retorno || '';
      sheet.getRangeByIndexes(row - 1, NotaCriada, 1, 1).setValue(nota);
      sheet.getRangeByIndexes(row - 1, RetornoAPI, 1, 1).setValue(retorno);
    }

    return { ok: true, updated: results.length };
  }

  return { error: 'unknown action' };
}

// Helper: cria a observação no mesmo formato do GAS original
function buildObservacao(row: any[]) {
  const ValorDocumento = formatCurrency(row[Valor]);
  const observacao = `${row[NomeDoServico]}(${row[Quantidade]} X ${ValorDocumento}) - ${row[Valor]}\n\n\n` +
    `${row[Descriminacao1]}\n\n${row[Descriminacao2]}\n\n Data Vencimento: ${row[VencimentoFatura]}`;
  return observacao;
}

function formatCurrency(value: any) {
  if (typeof value === 'number') {
    try {
      return value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    } catch (e) {
      return value.toString();
    }
  }
  return value;
}
