/**
 * Teste para simular falha na autenticação
 */

import { RealAPIClient } from '../src/api/real-api-client';

async function testAuthFailure() {
  console.log('❌ TESTE DE FALHA NA AUTENTICAÇÃO\n');
  console.log('=' .repeat(50));
  
  // Criar cliente com credenciais inválidas
  const apiClient = new RealAPIClient();
  
  // Simular falha modificando temporariamente a resposta
  const originalCallAPI = apiClient.callAPI;
  
  // Override para simular falha
  apiClient.callAPI = async function(endpoint: string, payload: any, context?: string) {
    console.log('🔴 Simulando falha na API...');
    
    if (endpoint.includes('/oauth/token')) {
      // Simular resposta de erro da API
      return {
        success: false,
        error: {
          error: 'invalid_grant',
          error_description: 'Invalid username or password',
          status: 401
        }
      };
    }
    
    return originalCallAPI.call(this, endpoint, payload, context);
  };
  
  try {
    console.log('🔐 Tentando autenticar com credenciais inválidas...\n');
    
    const result = await apiClient.authenticate();
    
    console.log('\n📊 RESULTADO:');
    console.log(`Sucesso: ${result.success}`);
    console.log(`Erro: ${result.error}`);
    
    if (!result.success) {
      console.log('\n✅ Falha capturada corretamente!');
      console.log('Detalhes do erro:', result.error);
    } else {
      console.log('\n❌ Erro: deveria ter falhado!');
    }
    
  } catch (error) {
    console.log('\n🔴 Exceção capturada:');
    console.log(error);
  }
  
  console.log('\n' + '=' .repeat(50));
  console.log('❌ Teste de falha concluído!');
}

testAuthFailure().catch(console.error);