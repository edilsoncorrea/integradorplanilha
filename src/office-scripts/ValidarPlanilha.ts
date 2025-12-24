// ValidarPlanilha.ts
// Conversão parcial de ValidarPlanilha.gs: cria descritores de consultas (GET) que precisam ser executadas
// por Power Automate e depois aplicadas de volta na planilha.

import {
  CodigoDaEmpresa, CodigoCliente, NomeDoCliente, IdentificadorCliente,
  CodigoDaOperacao, IdentificadorOperacao, CFOP, CodigoDoServico, IdentificadorServico,
  NomeDoServico, Quantidade, Valor, Descriminacao1, Descriminacao2, Codigoprazo,
  IdentificadorPrazo, FormaPagamentoEntrada, CodidoDaFormaDePagamento, IdentificadorFormaPagamento,
  DataEmissao, VencimentoFatura, NotaCriada, RetornoAPI
} from './Constantes';

export function main(workbook: any, inputs?: any): any {
  // action: 'buildQueries' | 'applyResults'
  const action = inputs && inputs.action ? inputs.action : 'buildQueries';

  if (action === 'buildQueries') {
    const sheet = workbook.getWorksheet ? workbook.getWorksheet('Documento') : null;
    // When running from Power Automate, workbook is not accessible directly here in file system.
    // Best practice: this script will be called as Office Script and will read the sheet.

    if (!sheet) return { error: 'Worksheet Documento not found in this execution context' };

    const used = sheet.getUsedRange();
    if (!used) return { queries: [] };
    const values = used.getValues() as any[][];

    const queries: any[] = [];

    for (let i = 2; i < values.length; i++) {
      const row = values[i];
      if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && (!row[IdentificadorCliente] || String(row[IdentificadorCliente]).trim() === '')) {
        // Need to query pessoa by CodigoCliente
        queries.push({ sheetRow: i + 1, method: 'GET', endpoint: `/api/pessoas/codigo/${row[CodigoCliente]}` });
      }
      // Similar checks for FormaPagamento, Operacao, Servico, Prazo
      if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && (!row[IdentificadorFormaPagamento] || String(row[IdentificadorFormaPagamento]).trim() === '')) {
        queries.push({ sheetRow: i + 1, method: 'GET', endpoint: `/api/formasPagamento` });
      }
      if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && (!row[IdentificadorOperacao] || String(row[IdentificadorOperacao]).trim() === '')) {
        queries.push({ sheetRow: i + 1, method: 'GET', endpoint: `/api/operacoes` });
      }
      if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && (!row[IdentificadorServico] || String(row[IdentificadorServico]).trim() === '')) {
        queries.push({ sheetRow: i + 1, method: 'GET', endpoint: `/api/produtos?codigo=${row[CodigoDoServico]}` });
      }
      if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && (!row[IdentificadorPrazo] || String(row[IdentificadorPrazo]).trim() === '')) {
        queries.push({ sheetRow: i + 1, method: 'GET', endpoint: `/api/prazos` });
      }
    }

    return { queries };
  }

  if (action === 'applyResults') {
    const results = inputs && inputs.results ? inputs.results : [];
    const sheet = workbook.getWorksheet('Documento');
    if (!sheet) return { error: 'Worksheet Documento not found' };

    for (const res of results) {
      const row = res.sheetRow;
      if (!row) continue;
      if (res.identificadorCliente) {
        sheet.getRangeByIndexes(row - 1, IdentificadorCliente, 1, 1).setValue(res.identificadorCliente);
      }
      if (res.nomeDoCliente) {
        sheet.getRangeByIndexes(row - 1, NomeDoCliente, 1, 1).setValue(res.nomeDoCliente);
      }
      if (res.identificadorFormaPagamento) {
        sheet.getRangeByIndexes(row - 1, IdentificadorFormaPagamento, 1, 1).setValue(res.identificadorFormaPagamento);
      }
      if (res.identificadorOperacao) {
        sheet.getRangeByIndexes(row - 1, IdentificadorOperacao, 1, 1).setValue(res.identificadorOperacao);
      }
      if (res.identificadorServico) {
        sheet.getRangeByIndexes(row - 1, IdentificadorServico, 1, 1).setValue(res.identificadorServico);
      }
      if (res.identificadorPrazo) {
        sheet.getRangeByIndexes(row - 1, IdentificadorPrazo, 1, 1).setValue(res.identificadorPrazo);
      }
      if (res.notaCriada) {
        sheet.getRangeByIndexes(row - 1, NotaCriada, 1, 1).setValue(res.notaCriada);
      }
      if (res.retorno) {
        sheet.getRangeByIndexes(row - 1, RetornoAPI, 1, 1).setValue(res.retorno);
      }
    }

    return { ok: true, updated: results.length };
  }

  return { error: 'unknown action' };
}
