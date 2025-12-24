/**
 * Teste para simular exceções na autenticação
 */

import { RealAPIClient } from '../src/api/real-api-client';

async function testException() {
  console.log('💥 TESTE DE EXCEÇÃO NA AUTENTICAÇÃO\n');
  console.log('=' .repeat(50));
  
  const apiClient = new RealAPIClient();
  
  // Override para simular exceção
  const originalCallAPI = apiClient.callAPI;
  apiClient.callAPI = async function(endpoint: string, payload: any, context?: string) {
    console.log('💥 Simulando exceção de rede...');
    
    if (endpoint.includes('/oauth/token')) {
      // Simular diferentes tipos de erro
      throw new Error('ECONNREFUSED: Connection refused');
    }
    
    return originalCallAPI.call(this, endpoint, payload, context);
  };
  
  try {
    console.log('🔐 Tentando autenticar (vai gerar exceção)...\n');
    
    const result = await apiClient.authenticate();
    
    console.log('\n📊 RESULTADO:');
    console.log(`Sucesso: ${result.success}`);
    console.log('Erro:', JSON.stringify(result.error, null, 2));
    
  } catch (error) {
    console.log('\n🔴 Exceção não capturada (não deveria acontecer):');
    console.log(error);
  }
  
  console.log('\n' + '=' .repeat(50));
  console.log('💥 Teste de exceção concluído!');
}

testException().catch(console.error);