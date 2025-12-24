/**
 * Teste com CSV alterado manualmente para "Não"
 */

import { CSVSimulator } from '../src/mocks/csv-simulator';
import { MockWorkbook, atualizarPlanilhaComResultado } from '../src/mocks/excel-mock';

function testeCSVAlterado() {
  console.log('📄 TESTE CSV ALTERADO MANUALMENTE\n');
  
  // 1. Mostrar estado atual do CSV
  console.log('📊 ESTADO ATUAL DO CSV:');
  CSVSimulator.mostrarEstado();
  
  // 2. Verificar quais linhas seriam processadas
  console.log('\n🔍 VERIFICANDO QUAIS LINHAS SERIAM PROCESSADAS:');
  const dadosCSV = CSVSimulator.carregarCSV();
  
  dadosCSV.slice(2).forEach((linha, index) => {
    const numeroLinha = index + 3;
    const notaCriada = linha[21];
    const retorno = linha[22];
    const jaProcessada = notaCriada === 'Sim';
    
    console.log(`  Linha ${numeroLinha}: NotaCriada="${notaCriada}", Retorno="${retorno}" → ${jaProcessada ? '❌ IGNORAR' : '✅ PROCESSAR'}`);
  });
  
  // 3. Simular processamento
  console.log('\n🔄 SIMULANDO PROCESSAMENTO...');
  const workbook = new MockWorkbook().loadRealData();
  
  // Processar apenas linhas que não foram processadas
  dadosCSV.slice(2).forEach((linha, index) => {
    const numeroLinha = index + 3;
    const notaCriada = linha[21];
    
    if (notaCriada !== 'Sim') {
      console.log(`📤 Processando linha ${numeroLinha}...`);
      atualizarPlanilhaComResultado(workbook, numeroLinha, true, `00A000${numeroLinha}${numeroLinha}${numeroLinha}`);
    }
  });
  
  // 4. Estado final
  console.log('\n📊 ESTADO FINAL:');
  CSVSimulator.mostrarEstado();
  
  console.log('\n✅ Teste concluído! Todas as linhas com "Não" foram processadas.');
}

testeCSVAlterado();