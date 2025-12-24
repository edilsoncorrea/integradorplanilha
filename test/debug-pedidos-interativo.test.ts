/**
 * Debug INTERATIVO - Pedidos de Venda com Breakpoints
 * Use Ctrl+Shift+D no VS Code e selecione "🚀 Debug Pedidos Interativo"
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { main as pedidoMain } from '../src/office-scripts/PedidoDeVenda';
import { RealAPIClient } from '../src/api/real-api-client';
import { apiDebugger } from '../src/debug/api-debugger';

async function debugPedidosInterativo() {
  console.log('🚀 DEBUG INTERATIVO - PEDIDOS DE VENDA');
  console.log('=' .repeat(60));
  
  // 🔴 BREAKPOINT 1: Carregamento de dados
  debugger; // <- Pare aqui para inspecionar
  console.log('\n📊 Carregando dados da planilha CSV...');
  
  const mockWorkbook = new MockWorkbook().loadRealData();
  const sheet = mockWorkbook.getWorksheet('Documento');
  const data = sheet?.getUsedRange()?.getValues();
  
  console.log(`✅ ${data?.length || 0} linhas carregadas`);
  
  // 🔴 BREAKPOINT 2: Geração de payloads
  debugger; // <- Pare aqui para ver dados carregados
  console.log('\n🏗️ Gerando payloads dos pedidos...');
  
  const result = pedidoMain(mockWorkbook, { 
    action: 'buildPedidoVendaFromSheet' 
  });
  
  console.log(`✅ ${result.payloads?.length || 0} payloads gerados`);
  
  if (result.payloads?.length > 0) {
    console.log('\n📦 Primeiro payload:');
    console.log(JSON.stringify(result.payloads[0].payload, null, 2));
  }
  
  // 🔴 BREAKPOINT 3: Autenticação
  debugger; // <- Pare aqui para inspecionar payloads
  console.log('\n🔐 Iniciando autenticação...');
  
  // Ativar breakpoints do API debugger
  apiDebugger.enableBreakpoint('AUTH_CHECK');
  apiDebugger.enableBreakpoint('BEFORE_API_CALL');
  apiDebugger.enableBreakpoint('AFTER_API_RESPONSE');
  
  const apiClient = new RealAPIClient();
  const authResult = await apiClient.authenticate();
  
  if (!authResult.success) {
    console.log(`❌ Falha na autenticação: ${authResult.error}`);
    return;
  }
  
  console.log(`✅ Autenticado! Token: ${authResult.token?.substring(0, 20)}...`);
  
  // 🔴 BREAKPOINT 4: Envio de pedidos
  debugger; // <- Pare aqui antes de enviar pedidos
  console.log('\n🚀 Enviando pedidos para API...');
  
  for (let i = 0; i < Math.min(result.payloads.length, 2); i++) {
    const pedido = result.payloads[i];
    
    console.log(`\n📤 Pedido ${i + 1} (linha ${pedido.sheetRow})`);
    console.log(`   Empresa: ${pedido.payload.CodigoEmpresa}`);
    console.log(`   Cliente: ${pedido.payload.IdentificadorCliente}`);
    console.log(`   Valor: ${pedido.payload.Itens[0].Valor}`);
    
    // 🔴 BREAKPOINT 5: Antes de cada pedido
    debugger; // <- Pare aqui para inspecionar cada pedido
    
    try {
      const pedidoResult = await apiClient.createPedido(pedido.payload);
      
      if (pedidoResult.success) {
        console.log(`   ✅ Sucesso: ${JSON.stringify(pedidoResult.data)}`);
      } else {
        console.log(`   ❌ Erro: ${JSON.stringify(pedidoResult.error)}`);
      }
      
    } catch (error) {
      console.log(`   ❌ Exceção: ${error}`);
    }
  }
  
  // 🔴 BREAKPOINT 6: Final
  debugger; // <- Pare aqui para ver resultados finais
  console.log('\n🎯 Debug interativo concluído!');
  console.log(`Total processado: ${result.payloads.length} pedidos`);
}

// Executar debug interativo
debugPedidosInterativo().catch(console.error);