/**
 * Teste interativo para criação de pedidos
 */

import { PedidoSimulator } from '../src/simulators/pedido-simulator';

function runInteractivePedidoTest() {
  console.log('🎮 Simulação Interativa - Criação de Pedidos\n');
  console.log('='.repeat(50));
  
  const simulator = new PedidoSimulator();
  
  // Executa fluxo completo
  const resultado = simulator.executarFluxoCompleto();
  
  console.log('\n' + '='.repeat(50));
  console.log('🎯 Simulação concluída!');
  
  if (resultado) {
    console.log(`\n📋 Processados: ${resultado.pedidos.length} pedidos`);
    console.log(`✅ Sucessos: ${resultado.sucessos}`);
    console.log(`❌ Erros: ${resultado.erros}`);
  }
}

runInteractivePedidoTest();