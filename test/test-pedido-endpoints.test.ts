/**
 * Teste para encontrar endpoint correto de pedidos
 */

import { RealAPIClient } from '../src/api/real-api-client';
import { PedidoSimulator } from '../src/simulators/pedido-simulator';

async function testPedidoEndpoints() {
  console.log('🔍 TESTANDO ENDPOINTS DE PEDIDOS\n');
  
  const apiClient = new RealAPIClient();
  
  // Autenticar primeiro
  const authResult = await apiClient.authenticate();
  if (!authResult.success) {
    console.log('❌ Falha na autenticação');
    return;
  }
  
  console.log('✅ Autenticado com sucesso\n');
  
  // Gerar um pedido dos dados CSV
  const simulator = new PedidoSimulator();
  const pedidos = simulator.gerarPedidosDaPlanilha();
  
  if (pedidos.length === 0) {
    console.log('❌ Nenhum pedido gerado');
    return;
  }
  
  const pedido = pedidos[0]; // Usar primeiro pedido
  console.log('📋 Testando com pedido da linha:', pedido.sheetRow);
  console.log('Cliente:', pedido.payload.IdentificadorCliente);
  console.log('Valor:', pedido.payload.Itens[0].Valor);
  
  // Endpoints possíveis para testar
  const endpoints = [
    '/api/PedidoVenda',
    '/PedidoVenda', 
    '/api/pedidovenda',
    '/pedidovenda',
    '/api/vendas',
    '/vendas',
    '/api/Pedido',
    '/Pedido'
  ];
  
  console.log('\n🚀 Testando endpoints...\n');
  
  for (const endpoint of endpoints) {
    console.log(`--- Testando: ${endpoint} ---`);
    
    try {
      const result = await apiClient.callAPI(endpoint, pedido.payload, `Teste ${endpoint}`);
      
      if (result.success) {
        console.log(`✅ SUCESSO! Endpoint encontrado: ${endpoint}`);
        console.log('Resposta:', JSON.stringify(result.data, null, 2));
        return endpoint; // Retorna o endpoint que funcionou
      } else {
        console.log(`❌ Status: ${(result.error as any)?.status || 'Erro'}`);
      }
    } catch (error) {
      console.log(`💥 Erro: ${error}`);
    }
    
    // Pausa entre testes
    await new Promise(resolve => setTimeout(resolve, 500));
  }
  
  console.log('\n❌ Nenhum endpoint funcionou');
  return null;
}

testPedidoEndpoints().catch(console.error);