/**
 * Debug para testar a função criarDocumento
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { criarDocumento, main } from '../src/office-scripts/DocumentoCompleto';

async function debugCriarDocumento() {
  console.log('📄 DEBUG - CRIAR DOCUMENTO\n');
  console.log('=' .repeat(60));
  
  // PASSO 1: Carregar dados simulados
  console.log('\n📊 PASSO 1: Carregando dados da planilha');
  const mockWorkbook = new MockWorkbook().loadRealData();
  const sheet = mockWorkbook.getWorksheet('Documento');
  
  if (!sheet) {
    console.log('❌ Planilha não encontrada');
    return;
  }
  
  const data = sheet.getUsedRange()?.getValues();
  console.log(`✅ Dados carregados: ${data?.length || 0} linhas`);
  
  // PASSO 2: Executar criarDocumento
  console.log('\n📄 PASSO 2: Executando criarDocumento()');
  
  try {
    const result = criarDocumento(mockWorkbook);
    
    // IMPORTANTE: Simular atualização do CSV para cada documento
    if (result.payloads && result.payloads.length > 0) {
      console.log('\n💾 Simulando atualização do CSV...');
      const { atualizarPlanilhaComResultado } = require('../src/mocks/excel-mock');
      
      result.payloads.forEach((doc: any, index: number) => {
        const identificadorSimulado = `00A000${String(index + 1).padStart(3, '0')}`;
        atualizarPlanilhaComResultado(mockWorkbook, doc.sheetRow, true, identificadorSimulado);
        console.log(`   ✅ Linha ${doc.sheetRow} atualizada: ${identificadorSimulado}`);
      });
    }
    
    console.log('\n📊 RESULTADO DA CRIAÇÃO:');
    console.log(`Sucesso: ${result.success}`);
    console.log(`Mensagem: ${result.message}`);
    console.log(`Documentos processados: ${result.processed}`);
    console.log(`Payloads gerados: ${result.payloads?.length || 0}`);
    
    // PASSO 3: Mostrar exemplo de payload
    if (result.payloads && result.payloads.length > 0) {
      console.log('\n📦 PASSO 3: Exemplo de payload (primeiro documento)');
      const primeiroPayload = result.payloads[0];
      
      console.log(`Linha da planilha: ${primeiroPayload.sheetRow}`);
      console.log(`Endpoint: ${primeiroPayload.endpoint}`);
      console.log('\nPayload completo:');
      console.log(JSON.stringify(primeiroPayload.payload, null, 2));
      
      // Verificar campos específicos do documento
      console.log('\n🔍 PASSO 4: Verificação de campos específicos');
      const payload = primeiroPayload.payload;
      console.log(`StatusNotaFiscalEletronica: ${payload.StatusNotaFiscalEletronica}`);
      console.log(`TipoDocumento: ${payload.TipoDocumento}`);
      console.log(`TipoPagamento: ${payload.TipoPagamento}`);
      console.log(`AliquotaConvenio: ${payload.Pagamentos[0]?.AliquotaConvenio}`);
      
      // Verificar observação
      console.log('\n📝 Observação gerada:');
      console.log(payload.Observacao);
    }
    
    // PASSO 5: Testar via main() também
    console.log('\n🔧 PASSO 5: Testando via main()');
    const mainResult = main(mockWorkbook, { action: 'criarDocumento' });
    
    console.log('Resultado via main():');
    console.log(`Sucesso: ${mainResult.success}`);
    console.log(`Mensagem: ${mainResult.message}`);
    console.log(`Processados: ${mainResult.processed}`);
    
  } catch (error) {
    console.log(`\n❌ ERRO: ${error}`);
  }
  
  console.log('\n' + '=' .repeat(60));
  console.log('🎯 Debug criarDocumento concluído!');
  
  console.log('\n💡 Diferenças vs Pedidos de Venda:');
  console.log('   📄 Documentos: /api/documentos (notas fiscais)');
  console.log('   📦 Pedidos: /api/venda/pedidos (pedidos de venda)');
  console.log('   🔧 Campo extra: AliquotaConvenio nos documentos');
}

debugCriarDocumento().catch(console.error);