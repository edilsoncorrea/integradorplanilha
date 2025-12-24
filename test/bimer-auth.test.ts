/**
 * Teste específico para autenticação Bimer
 */

import { main } from '../src/office-scripts/Autenticacao';
import { RealAPIClient } from '../src/api/real-api-client';

async function testBimerAuth() {
  console.log('🔐 TESTE DE AUTENTICAÇÃO BIMER\n');
  console.log('=' .repeat(50));
  
  // 1. Testar função de autenticação diretamente
  console.log('📋 Testando função buildAuthPayload...');
  
  const authData = main(null as any, {
    action: 'buildAuthPayload',
    host: 'https://087344bimerapi.alterdata.cloud',
    username: 'bimerapi',
    senha: '123456',
    nonce: '123456789'
  });
  
  console.log('✅ Payload de autenticação gerado:');
  console.log(`   URL: ${authData.url}`);
  console.log(`   Method: ${authData.method}`);
  console.log('   Payload:');
  console.log(JSON.stringify(authData.payload, null, 4));
  
  // 2. Testar hash MD5
  console.log('\n🔒 Testando hash MD5...');
  const hashTest = main(null as any, {
    action: 'hash',
    value: 'bimerapi123456789123456'
  });
  
  console.log(`✅ Hash MD5: ${hashTest.md5}`);
  
  // 3. Testar com cliente da API
  console.log('\n🌐 Testando com cliente da API...');
  const apiClient = new RealAPIClient();
  
  try {
    const result = await apiClient.authenticate();
    
    if (result.success) {
      console.log('✅ Autenticação simulada bem-sucedida');
      console.log(`   Token: ${result.token?.substring(0, 20)}...`);
    } else {
      console.log('❌ Falha na autenticação simulada');
    }
  } catch (error) {
    console.log('❌ Erro:', error);
  }
  
  console.log('\n' + '=' .repeat(50));
  console.log('🔐 Teste de autenticação concluído!');
}

testBimerAuth().catch(console.error);