/**
 * Teste com API real e debug completo
 */

import { RealAPIClient } from '../src/api/real-api-client';
import { PedidoSimulator } from '../src/simulators/pedido-simulator';
import { apiDebugger } from '../src/debug/api-debugger';

async function runRealAPIDebugTest() {
  console.log('🔍 TESTE DE DEBUG - API REAL\n');
  console.log('=' .repeat(60));
  
  // 1. Configurar cliente da API
  const apiClient = new RealAPIClient();
  
  // 2. Autenticar automaticamente com dados hardcoded
  console.log('\n🔐 Autenticando com API Bimer...');
  const authResult = await apiClient.authenticate();
  
  if (!authResult.success) {
    console.log('❌ Falha na autenticação. Usando token simulado.');
    const fakeToken = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.exemplo';
    apiClient.setAuthToken(fakeToken);
  }
  
  // 3. Testar conectividade
  console.log('\n🔌 Testando conectividade...');
  await apiClient.testConnection();
  
  // 4. Gerar pedidos da planilha
  console.log('\n📋 Gerando pedidos da planilha...');
  const simulator = new PedidoSimulator();
  const pedidos = simulator.gerarPedidosDaPlanilha();
  
  // 5. Enviar cada pedido para API real
  console.log('\n🚀 Enviando pedidos para API real...');
  
  for (let i = 0; i < pedidos.length; i++) {
    const pedido = pedidos[i];
    console.log(`\n--- PEDIDO ${i + 1} ---`);
    
    const result = await apiClient.callAPI(
      '/api/pedidos', 
      pedido.payload, 
      `Pedido ${i + 1} - Linha ${pedido.sheetRow}`
    );
    
    console.log(`Resultado: ${result.success ? 'SUCESSO' : 'ERRO'}`);
    if (result.data) {
      console.log(`ID: ${(result.data as any)?.id || 'N/A'}`);
    }
  }
  
  console.log('\n' + '=' .repeat(60));
  console.log('🔍 Debug concluído!');
  
  console.log('\n💡 Para usar API real:');
  console.log('   1. Descomente o código fetch() em real-api-client.ts');
  console.log('   2. A autenticação usa dados hardcoded do arquivo Autenticação.ts');
  console.log('   3. Credenciais: bimerapi / 123456 / nonce: 123456789');
}

runRealAPIDebugTest().catch(console.error);