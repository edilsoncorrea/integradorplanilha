// DocumentoCompleto.ts
// Conversão completa de Documento.gs para Office Scripts
// Inclui a função principal criarDocumento() e main() com actions

import {
  CodigoDaEmpresa, IdentificadorOperacao, IdentificadorCliente, NomeDoServico,
  Quantidade, Valor, Descriminacao1, Descriminacao2, VencimentoFatura, DataEmissao,
  CFOP, IdentificadorServico, IdentificadorFormaPagamento, NotaCriada, RetornoAPI
} from './Constantes';

// Função principal - equivale ao criarDocumento() do .gs
export function criarDocumento(workbook: any): any {
  console.log('📄 Iniciando criação de documentos...');
  
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };
  
  const used = sheet.getUsedRange();
  if (!used) return { processed: 0 };
  const values = used.getValues() as any[][];
  
  let processed = 0;
  const results = [];
  
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const notaCriada = String(row[NotaCriada] || '').trim();
    if (notaCriada === '' || notaCriada === 'Não') {
      
      // Extrair valor numérico da string (igual aos pedidos)
      const valorString = String(row[Valor]);
      const valorNumerico = parseFloat(valorString.replace(/[R$\s.,]/g, '').replace(',', '.')) || 0;
      
      // Criar observação rica como no .gs
      const valorDocumento = formatCurrency(valorNumerico);
      const observacao = 
        row[NomeDoServico] + '(' + row[Quantidade] + ' X ' + valorDocumento + ') - ' + Valor + '\n\n\n' +
        row[Descriminacao1] + '\n\n' + 
        row[Descriminacao2] + 
        '\n\n Data Vencimento: ' + 
        row[VencimentoFatura];
      
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
            ValorUnitario: valorNumerico  // ← Corrigido: usar valor numérico
          }
        ],
        Pagamentos: [
          {
            Aliquota: 100,
            AliquotaConvenio: 10,
            DataVencimento: row[DataEmissao],
            IdentificadorFormaPagamento: row[IdentificadorFormaPagamento],
            Valor: valorNumerico  // ← Corrigido: usar valor numérico
          }
        ]
      };
      
      results.push({
        sheetRow: i + 1,
        payload,
        endpoint: '/api/documentos'
      });
      
      processed++;
    }
  }
  
  console.log(`📄 ${processed} documentos preparados para criação`);
  
  return {
    success: true,
    message: `${processed} documentos preparados`,
    payloads: results,
    processed
  };
}

export function main(workbook: any, inputs?: any): any {
  const action = inputs?.action || 'criarDocumento';
  
  if (action === 'criarDocumento') {
    return criarDocumento(workbook);
  }

  if (action === 'buildDocumentoPayloads') {
    const sheet = workbook.getWorksheet('Documento');
    if (!sheet) return { error: 'Worksheet Documento not found' };
    const used = sheet.getUsedRange();
    if (!used) return { payloads: [] };
    const values = used.getValues() as any[][];

    const payloads: any[] = [];
    for (let i = 2; i < values.length; i++) {
      const row = values[i];
      const notaCriada = String(row[NotaCriada] || '').trim();
      if (notaCriada === '' || notaCriada === 'Não') {
        
        // Extrair valor numérico (igual à função principal)
        const valorString = String(row[Valor]);
        const valorNumerico = parseFloat(valorString.replace(/[R$\s.,]/g, '').replace(',', '.')) || 0;
        
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
              ValorUnitario: valorNumerico  // ← Corrigido: usar valor numérico
            }
          ],
          Pagamentos: [
            {
              Aliquota: 100,
              AliquotaConvenio: 10,
              DataVencimento: row[DataEmissao],
              IdentificadorFormaPagamento: row[IdentificadorFormaPagamento],
              Valor: valorNumerico  // ← Corrigido: usar valor numérico
            }
          ]
        };

        payloads.push({ sheetRow: i + 1, payload });
      }
    }

    return { payloads };
  }

  if (action === 'applyDocumentoResults') {
    const results = inputs?.results || [];
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