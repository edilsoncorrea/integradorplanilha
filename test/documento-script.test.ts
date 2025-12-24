/**
 * Testes para DocumentoScript.ts
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { main } from '../src/office-scripts/DocumentoScript';

// Simula o ambiente global ExcelScript
(global as any).ExcelScript = {};

function runTests() {
  console.log('🧪 Iniciando testes do DocumentoScript...\n');

  // Teste 1: buildPayloads
  console.log('📋 Teste 1: buildPayloads');
  const mockWorkbook = new MockWorkbook();
  
  const result = main(mockWorkbook as any, { action: 'buildPayloads' });
  
  console.log('✅ Payloads gerados:', result.payloads?.length || 0);
  if (result.payloads && result.payloads.length > 0) {
    console.log('📄 Primeiro payload:');
    console.log(JSON.stringify(result.payloads[0], null, 2));
  }
  console.log('');

  // Teste 2: applyResults
  console.log('📝 Teste 2: applyResults');
  const testResults = [
    { sheetRow: 3, notaCriada: 'NF001', retorno: 'Sucesso' },
    { sheetRow: 4, notaCriada: 'NF002', retorno: 'Sucesso' }
  ];
  
  const applyResult = main(mockWorkbook as any, { 
    action: 'applyResults', 
    results: testResults 
  });
  
  console.log('✅ Resultados aplicados:', applyResult);
  console.log('');

  // Teste 3: Validação de dados
  console.log('🔍 Teste 3: Validação de estrutura de payload');
  if (result.payloads && result.payloads.length > 0) {
    const payload = result.payloads[0].payload;
    const requiredFields = [
      'StatusNotaFiscalEletronica',
      'TipoDocumento', 
      'CodigoEmpresa',
      'DataEmissao',
      'IdentificadorOperacao',
      'IdentificadorPessoa',
      'Itens',
      'Pagamentos'
    ];
    
    const missingFields = requiredFields.filter(field => !(field in payload));
    
    if (missingFields.length === 0) {
      console.log('✅ Todos os campos obrigatórios estão presentes');
    } else {
      console.log('❌ Campos faltando:', missingFields);
    }
    
    // Validar estrutura dos itens
    if (payload.Itens && payload.Itens.length > 0) {
      const item = payload.Itens[0];
      const itemFields = ['CFOP', 'IdentificadorProduto', 'Quantidade', 'ValorUnitario'];
      const missingItemFields = itemFields.filter(field => !(field in item));
      
      if (missingItemFields.length === 0) {
        console.log('✅ Estrutura dos itens está correta');
      } else {
        console.log('❌ Campos faltando nos itens:', missingItemFields);
      }
    }
  }
  console.log('');

  // Teste 4: Cenário sem dados
  console.log('🚫 Teste 4: Planilha vazia');
  const emptyWorkbook = new MockWorkbook();
  emptyWorkbook.addWorksheet('Documento', []);
  
  const emptyResult = main(emptyWorkbook as any, { action: 'buildPayloads' });
  console.log('✅ Resultado com planilha vazia:', emptyResult);
  console.log('');

  console.log('🎉 Todos os testes concluídos!');
}

// Executa os testes
runTests();