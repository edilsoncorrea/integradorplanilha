/**
 * Teste com breakpoints de API ativados
 */

import { PedidoSimulator } from '../src/simulators/pedido-simulator';
import { apiDebugger } from '../src/debug/api-debugger';

function runDebugAPITest() {
  console.log('🐛 Teste com Breakpoints de API\n');
  console.log('=' .repeat(60));
  
  // Mostra status dos breakpoints
  apiDebugger.listBreakpoints();
  
  console.log('\n🎯 Iniciando simulação com breakpoints...\n');
  
  const simulator = new PedidoSimulator();
  
  // Executa com breakpoints ativos
  const resultado = simulator.executarFluxoCompleto();
  
  console.log('\n' + '=' .repeat(60));
  console.log('🐛 Debug concluído!');
  
  if (resultado) {
    console.log(`\n📊 Estatísticas finais:`);
    console.log(`   📋 Total processado: ${resultado.pedidos.length}`);
    console.log(`   ✅ Sucessos: ${resultado.sucessos}`);
    console.log(`   ❌ Erros: ${resultado.erros}`);
    console.log(`   📈 Taxa de sucesso: ${((resultado.sucessos / resultado.pedidos.length) * 100).toFixed(1)}%`);
  }
  
  console.log('\n💡 Dicas para debug real:');
  console.log('   • Use debugger; nos pontos críticos');
  console.log('   • Configure breakpoints no VS Code');
  console.log('   • Inspecione variáveis no momento da pausa');
}

runDebugAPITest();