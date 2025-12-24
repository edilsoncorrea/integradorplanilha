/**
 * Testes com dados reais da planilha CSV
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { main } from '../src/office-scripts/DocumentoScript';

(global as any).ExcelScript = {};

function runRealDataTests() {
  console.log('📊 Testando com dados reais da planilha...\n');

  const mockWorkbook = new MockWorkbook();
  mockWorkbook.loadRealData();

  // Teste 1: Verificar dados carregados
  console.log('📋 Teste 1: Verificação dos dados');
  const sheet = mockWorkbook.getWorksheet('Documento');
  const data = sheet?.getUsedRange()?.getValues();
  console.log(`✅ Linhas carregadas: ${data?.length || 0}`);
  console.log(`✅ Colunas: ${data?.[0]?.length || 0}`);
  console.log('');

  // Teste 2: buildPayloads com dados reais
  console.log('🔧 Teste 2: buildPayloads com dados reais');
  const result = main(mockWorkbook as any, { action: 'buildPayloads' });
  
  console.log(`✅ Payloads gerados: ${result.payloads?.length || 0}`);
  
  if (result.payloads && result.payloads.length > 0) {
    console.log('📄 Primeiro payload (linha com "Nota Criada?" vazia):');
    const firstPayload = result.payloads.find((p: any) => p.sheetRow === 5); // Última linha sem nota
    if (firstPayload) {
      console.log(`   Linha: ${firstPayload.sheetRow}`);
      console.log(`   Empresa: ${firstPayload.payload.CodigoEmpresa}`);
      console.log(`   Cliente: ${firstPayload.payload.IdentificadorPessoa}`);
      console.log(`   Serviço: ${firstPayload.payload.Itens[0].IdentificadorProduto}`);
      console.log(`   Valor: ${firstPayload.payload.Itens[0].ValorUnitario}`);
    }
  }
  console.log('');

  // Teste 3: Validar cenários específicos
  console.log('🎯 Teste 3: Cenários específicos');
  
  // Verificar se linhas com "Sim" em "Nota Criada?" são ignoradas
  const linesWithNota = data?.filter((row, index) => 
    index > 1 && row[21] === 'Sim'
  ).length || 0;
  
  const expectedPayloads = (data?.length || 0) - 2 - linesWithNota; // -2 para headers, -linesWithNota
  
  console.log(`✅ Linhas com nota criada (ignoradas): ${linesWithNota}`);
  console.log(`✅ Payloads esperados: ${expectedPayloads}`);
  console.log(`✅ Payloads gerados: ${result.payloads?.length || 0}`);
  console.log(`${expectedPayloads === (result.payloads?.length || 0) ? '✅' : '❌'} Validação: ${expectedPayloads === (result.payloads?.length || 0) ? 'PASSOU' : 'FALHOU'}`);
  console.log('');

  // Teste 4: Aplicar resultados
  console.log('📝 Teste 4: Aplicar resultados');
  const testResults = [
    { sheetRow: 5, notaCriada: 'NF003', retorno: 'Sucesso - Teste' }
  ];
  
  const applyResult = main(mockWorkbook as any, { 
    action: 'applyResults', 
    results: testResults 
  });
  
  console.log('✅ Resultado da aplicação:', applyResult);
  console.log('');

  console.log('🎉 Testes com dados reais concluídos!');
}

runRealDataTests();