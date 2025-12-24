/**
 * Debug passo a passo da autenticação
 */

import { main } from '../src/office-scripts/Autenticacao';
import { RealAPIClient } from '../src/api/real-api-client';
import { apiDebugger } from '../src/debug/api-debugger';

async function debugAuthStepByStep() {
  console.log('🔍 DEBUG PASSO A PASSO - AUTENTICAÇÃO\n');
  console.log('=' .repeat(60));
  
  // PASSO 1: Verificar dados de entrada
  console.log('\n📋 PASSO 1: Dados de entrada');
  const credentials = {
    username: 'bimerapi',
    senha: '123456', 
    nonce: '123456789'
  };
  
  console.log(`Username: ${credentials.username}`);
  console.log(`Senha: ${credentials.senha}`);
  console.log(`Nonce: ${credentials.nonce}`);
  
  // PASSO 2: Gerar hash MD5
  console.log('\n🔒 PASSO 2: Geração do hash MD5');
  const inputString = credentials.username + credentials.nonce + credentials.senha;
  console.log(`String para hash: "${inputString}"`);
  
  const hashResult = main(null as any, {
    action: 'hash',
    value: inputString
  });
  
  console.log(`Hash MD5 gerado: ${hashResult.md5}`);
  
  // PASSO 3: Construir payload de autenticação
  console.log('\n📦 PASSO 3: Payload de autenticação');
  const authPayload = main(null as any, {
    action: 'buildAuthPayload',
    host: 'https://087344bimerapi.alterdata.cloud',
    username: credentials.username,
    senha: credentials.senha,
    nonce: credentials.nonce
  });
  
  console.log('Payload completo:');
  console.log(JSON.stringify(authPayload, null, 2));
  
  // PASSO 4: Simular chamada HTTP
  console.log('\n🌐 PASSO 4: Simulação de chamada HTTP');
  console.log(`Endpoint: ${authPayload.url}`);
  console.log(`Método: ${authPayload.method}`);
  console.log('Headers esperados:');
  console.log('  Content-Type: application/json');
  console.log('  Accept: application/json');
  
  // PASSO 5: Teste com cliente real
  console.log('\n🔧 PASSO 5: Teste com cliente da API');
  
  // Ativar todos os breakpoints
  apiDebugger.enableBreakpoint('AUTH_CHECK');
  apiDebugger.enableBreakpoint('BEFORE_API_CALL');
  apiDebugger.enableBreakpoint('AFTER_API_RESPONSE');
  
  const apiClient = new RealAPIClient();
  
  try {
    const result = await apiClient.authenticate();
    
    console.log('\n📊 RESULTADO FINAL:');
    console.log(`Sucesso: ${result.success}`);
    if (result.success) {
      console.log(`Token: ${result.token}`);
    } else {
      console.log(`Erro: ${result.error}`);
    }
    
  } catch (error) {
    console.log('\n❌ ERRO CAPTURADO:');
    console.log(error);
  }
  
  console.log('\n' + '=' .repeat(60));
  console.log('🔍 Debug passo a passo concluído!');
  
  console.log('\n💡 Para API real:');
  console.log('   1. Descomente fetch() em real-api-client.ts');
  console.log('   2. Verifique se o endpoint /oauth/token está correto');
  console.log('   3. Confirme se as credenciais estão válidas');
}

debugAuthStepByStep().catch(console.error);