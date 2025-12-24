/**
 * Teste para mostrar o vetor realData atualizado como código
 */

import { MockWorkbook, atualizarPlanilhaComResultado, mostrarVetorAtualizado } from '../src/mocks/excel-mock';

function testeVetorCodigo() {
  console.log('📋 MOSTRAR VETOR ATUALIZADO COMO CÓDIGO\n');
  
  // 1. Criar workbook
  const workbook = new MockWorkbook().loadRealData();
  
  // 2. Simular algumas atualizações
  console.log('🔄 Simulando processamento...');
  atualizarPlanilhaComResultado(workbook, 3, true, '00A0000111');
  atualizarPlanilhaComResultado(workbook, 4, true, '00A0000222');
  
  // 3. Mostrar vetor atualizado como código
  mostrarVetorAtualizado(workbook);
  
  console.log('\n✅ Agora você pode copiar o código acima e colar no loadRealData() se quiser!');
}

testeVetorCodigo();