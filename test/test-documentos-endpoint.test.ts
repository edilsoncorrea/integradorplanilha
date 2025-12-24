/**
 * Teste com endpoint /api/documentos encontrado
 */

import { RealAPIClient } from '../src/api/real-api-client';
import { PedidoSimulator } from '../src/simulators/pedido-simulator';

async function testDocumentosEndpoint() {
  console.log('📄 TESTANDO ENDPOINT /api/documentos\n');
  
  const apiClient = new RealAPIClient();
  
  // Autenticar
  const authResult = await apiClient.authenticate();
  if (!authResult.success) {
    console.log('❌ Falha na autenticação');
    return;
  }
  
  console.log('✅ Autenticado\n');
  
  // Gerar pedido dos dados CSV
  const simulator = new PedidoSimulator();
  const pedidos = simulator.gerarPedidosDaPlanilha();
  
  if (pedidos.length === 0) {
    console.log('❌ Nenhum pedido gerado');
    return;
  }
  
  const pedido = pedidos[0];
  console.log('📋 Testando com pedido da linha:', pedido.sheetRow);
  console.log('Cliente:', pedido.payload.IdentificadorCliente);
  console.log('Valor:', pedido.payload.Itens[0].Valor);
  
  console.log('\n🚀 Enviando para /api/documentos...\n');
  
  try {
    const result = await apiClient.callAPI('/api/documentos', pedido.payload, 'Teste Documentos');
    
    if (result.success) {
      console.log('✅ SUCESSO!');
      console.log('Resposta:', JSON.stringify(result.data, null, 2));
    } else {
      console.log('❌ Erro:', result.error);
      
      // Se for erro 500, mostrar detalhes
      if (typeof result.error === 'string' && result.error.includes('500')) {
        console.log('\n🔍 Erro 500 indica problema no payload. Vamos ajustar...');
      }
    }
    
  } catch (error) {
    console.log('💥 Exceção:', error);
  }
}

testDocumentosEndpoint().catch(console.error);