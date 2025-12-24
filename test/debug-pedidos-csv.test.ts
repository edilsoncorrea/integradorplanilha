/**
 * Debug completo: CSV → Payloads → API Bimer
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { main as pedidoMain } from '../src/office-scripts/PedidoDeVenda';
import { main as authMain } from '../src/office-scripts/Autenticacao';
import { RealAPIClient } from '../src/api/real-api-client';

async function debugPedidosComCSV() {
  console.log('🔍 DEBUG COMPLETO - PEDIDOS DE VENDA COM CSV\n');
  console.log('=' .repeat(70));
  
  // PASSO 1: Carregar dados do CSV (via mock)
  console.log('\n📊 PASSO 1: Carregando dados da planilha');
  const mockWorkbook = new MockWorkbook().loadRealData();
  const sheet = mockWorkbook.getWorksheet('Documento');
  
  if (!sheet) {
    console.log('❌ Planilha não encontrada');
    return;
  }
  
  const data = sheet.getUsedRange()?.getValues();
  console.log(`✅ Dados carregados: ${data?.length || 0} linhas`);
  
  // Mostrar headers
  if (data && data.length > 1) {
    console.log('\n📋 Headers da planilha:');
    data[1].forEach((header: string, index: number) => {
      if (header) console.log(`  [${index}] ${header}`);
    });
  }
  
  // PASSO 2: Gerar payloads dos pedidos
  console.log('\n🏗️ PASSO 2: Gerando payloads dos pedidos');
  const result = pedidoMain(mockWorkbook, { action: 'buildPedidoVendaFromSheet' });
  
  if (result.error) {
    console.log(`❌ Erro: ${result.error}`);
    return;
  }
  
  console.log(`✅ ${result.payloads.length} payloads gerados`);
  
  // Mostrar primeiro payload como exemplo
  if (result.payloads.length > 0) {
    console.log('\n📦 Exemplo de payload (primeiro pedido):');
    console.log(JSON.stringify(result.payloads[0].payload, null, 2));
  }
  
  // PASSO 3: Autenticar na API
  console.log('\n🔐 PASSO 3: Autenticando na API Bimer');
  const apiClient = new RealAPIClient();
  
  try {
    const authResult = await apiClient.authenticate();
    
    if (!authResult.success) {
      console.log(`❌ Falha na autenticação: ${authResult.error}`);
      console.log('\n💡 Dica: Verifique se o fetch() está descomentado em real-api-client.ts');
      return;
    }
    
    console.log(`✅ Autenticado com sucesso!`);
    console.log(`🎫 Token: ${authResult.token?.substring(0, 20)}...`);
    
    // PASSO 4: Enviar pedidos para API
    console.log('\n🚀 PASSO 4: Enviando pedidos para API');
    
    for (let i = 0; i < Math.min(result.payloads.length, 2); i++) {
      const pedido = result.payloads[i];
      console.log(`\n📤 Enviando pedido ${i + 1} (linha ${pedido.sheetRow})`);
      
      try {
        // Simular envio do pedido
        console.log(`   Empresa: ${pedido.payload.CodigoEmpresa}`);
        console.log(`   Cliente: ${pedido.payload.IdentificadorCliente}`);
        console.log(`   Valor: R$ ${pedido.payload.Itens[0].Valor}`);
        
        // Fazer chamada real para API (descomente para testar)
        const pedidoResult = await apiClient.createPedido(pedido.payload);
        console.log(`   📋 Resultado: ${JSON.stringify(pedidoResult, null, 2)}`);
        
        console.log(`   ✅ Pedido processado (simulado - descomente acima para API real)`);
        
      } catch (error) {
        console.log(`   ❌ Erro no pedido: ${error}`);
      }
    }
    
  } catch (error) {
    console.log(`❌ Erro na autenticação: ${error}`);
  }
  
  // PASSO 5: Resumo dos dados processados
  console.log('\n📊 PASSO 5: Resumo dos dados');
  console.log(`Total de linhas na planilha: ${data?.length || 0}`);
  console.log(`Pedidos para processar: ${result.payloads.length}`);
  console.log(`Pedidos já criados: ${data?.filter((row: any[], i: number) => 
    i > 1 && row[21] && row[21].toString().trim() !== '').length || 0}`);
  
  console.log('\n' + '=' .repeat(70));
  console.log('🎯 Debug completo finalizado!');
  
  console.log('\n💡 Próximos passos:');
  console.log('   1. Para API real: descomente fetch() em real-api-client.ts');
  console.log('   2. Para envio real: descomente apiClient.createPedido()');
  console.log('   3. Verifique credenciais da API se houver erro de auth');
}

// Executar debug
debugPedidosComCSV().catch(console.error);