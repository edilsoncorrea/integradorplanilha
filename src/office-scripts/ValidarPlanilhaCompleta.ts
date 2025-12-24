// ValidarPlanilhaCompleta.ts
// Conversão completa de ValidarPlanilha.gs para Office Scripts
// Inclui a função principal ValidarPlanilha() e todas as subfunções

import {
  CodigoDaEmpresa, CodigoCliente, NomeDoCliente, IdentificadorCliente,
  CodigoDaOperacao, IdentificadorOperacao, CFOP, CodigoDoServico, IdentificadorServico,
  NomeDoServico, Quantidade, Valor, Descriminacao1, Descriminacao2, Codigoprazo,
  IdentificadorPrazo, FormaPagamentoEntrada, CodidoDaFormaDePagamento, IdentificadorFormaPagamento,
  DataEmissao, VencimentoFatura, NotaCriada, RetornoAPI
} from './Constantes';

// Função principal - equivale ao ValidarPlanilha() do .gs
export function ValidarPlanilha(workbook: any): any {
  console.log('🔍 Iniciando validação da planilha...');
  
  const results = {
    consultarCliente: ConsultarCliente(workbook),
    consultarFormaPagamento: ConsultarFormaPagamento(workbook),
    consultarOperacao: ConsultarOperacao(workbook),
    consultarServico: ConsultarServico(workbook),
    consultarPrazo: ConsultarPrazo(workbook)
  };
  
  const totalProcessed = Object.values(results).reduce((sum: number, result: any) => 
    sum + (result.processed || 0), 0);
  
  return {
    success: true,
    message: `Validação concluída. ${totalProcessed} consultas processadas.`,
    results
  };
}

export function main(workbook: any, inputs?: any): any {
  const action = inputs?.action || 'ValidarPlanilha';
  
  if (action === 'ValidarPlanilha') {
    return ValidarPlanilha(workbook);
  }
  
  // Manter compatibilidade com ações existentes
  if (action === 'buildQueries') {
    const sheet = workbook.getWorksheet('Documento');
    if (!sheet) return { error: 'Worksheet Documento not found' };
    
    const used = sheet.getUsedRange();
    if (!used) return { queries: [] };
    const values = used.getValues() as any[][];
    
    const queries: any[] = [];
    
    for (let i = 2; i < values.length; i++) {
      const row = values[i];
      if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && (!row[IdentificadorCliente] || String(row[IdentificadorCliente]).trim() === '')) {
        queries.push({ sheetRow: i + 1, method: 'GET', endpoint: `/api/pessoas/codigo/${row[CodigoCliente]}` });
      }
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
  
  return { error: 'unknown action' };
}

// ConsultarCliente - equivale ao ConsultarCliente() do .gs
function ConsultarCliente(workbook: any): any {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };
  
  const used = sheet.getUsedRange();
  if (!used) return { processed: 0 };
  const values = used.getValues() as any[][];
  
  let processed = 0;
  const queries = [];
  
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && 
        (!row[IdentificadorCliente] || String(row[IdentificadorCliente]).trim() === '')) {
      
      queries.push({
        sheetRow: i + 1,
        endpoint: `/api/pessoas/codigo/${row[CodigoCliente]}`,
        codigoCliente: row[CodigoCliente],
        type: 'cliente'
      });
      processed++;
    }
  }
  
  console.log(`📋 ConsultarCliente: ${processed} consultas preparadas`);
  return { processed, queries, type: 'cliente' };
}

// ConsultarFormaPagamento - equivale ao ConsultarFormaPagamento() do .gs
function ConsultarFormaPagamento(workbook: any): any {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };
  
  const used = sheet.getUsedRange();
  if (!used) return { processed: 0 };
  const values = used.getValues() as any[][];
  
  let processed = 0;
  const queries = [];
  
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && 
        (!row[IdentificadorFormaPagamento] || String(row[IdentificadorFormaPagamento]).trim() === '')) {
      
      queries.push({
        sheetRow: i + 1,
        endpoint: '/api/formasPagamento',
        codigoFormaPagamento: row[CodidoDaFormaDePagamento],
        type: 'formaPagamento'
      });
      processed++;
    }
  }
  
  console.log(`💳 ConsultarFormaPagamento: ${processed} consultas preparadas`);
  return { processed, queries, type: 'formaPagamento' };
}

// ConsultarOperacao - equivale ao ConsultarOperacao() do .gs
function ConsultarOperacao(workbook: any): any {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };
  
  const used = sheet.getUsedRange();
  if (!used) return { processed: 0 };
  const values = used.getValues() as any[][];
  
  let processed = 0;
  const queries = [];
  
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && 
        (!row[IdentificadorOperacao] || String(row[IdentificadorOperacao]).trim() === '')) {
      
      queries.push({
        sheetRow: i + 1,
        endpoint: '/api/operacoes',
        codigoOperacao: row[CodigoDaOperacao],
        type: 'operacao'
      });
      processed++;
    }
  }
  
  console.log(`⚙️ ConsultarOperacao: ${processed} consultas preparadas`);
  return { processed, queries, type: 'operacao' };
}

// ConsultarServico - equivale ao ConsultarServico() do .gs
function ConsultarServico(workbook: any): any {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };
  
  const used = sheet.getUsedRange();
  if (!used) return { processed: 0 };
  const values = used.getValues() as any[][];
  
  let processed = 0;
  const queries = [];
  
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && 
        (!row[IdentificadorServico] || String(row[IdentificadorServico]).trim() === '')) {
      
      queries.push({
        sheetRow: i + 1,
        endpoint: `/api/produtos?codigo=${row[CodigoDoServico]}`,
        codigoServico: row[CodigoDoServico],
        type: 'servico'
      });
      processed++;
    }
  }
  
  console.log(`🛠️ ConsultarServico: ${processed} consultas preparadas`);
  return { processed, queries, type: 'servico' };
}

// ConsultarPrazo - equivale ao ConsultarPrazo() do .gs
function ConsultarPrazo(workbook: any): any {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Worksheet Documento not found' };
  
  const used = sheet.getUsedRange();
  if (!used) return { processed: 0 };
  const values = used.getValues() as any[][];
  
  let processed = 0;
  const queries = [];
  
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    if ((!row[NotaCriada] || String(row[NotaCriada]).trim() === '') && 
        (!row[IdentificadorPrazo] || String(row[IdentificadorPrazo]).trim() === '')) {
      
      queries.push({
        sheetRow: i + 1,
        endpoint: '/api/prazos',
        codigoPrazo: row[Codigoprazo],
        type: 'prazo'
      });
      processed++;
    }
  }
  
  console.log(`📅 ConsultarPrazo: ${processed} consultas preparadas`);
  return { processed, queries, type: 'prazo' };
}