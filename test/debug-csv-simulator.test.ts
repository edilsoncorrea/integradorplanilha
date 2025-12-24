/**
 * Teste do CSV Simulator - Simulação completa da planilha
 */

import { CSVSimulator } from '../src/mocks/csv-simulator';
import { MockWorkbook, atualizarPlanilhaComResultado } from '../src/mocks/excel-mock';

function testeCSVSimulator() {
  console.log('📄 TESTE CSV SIMULATOR - SIMULAÇÃO COMPLETA DA PLANILHA\n');
  
  // 1. Resetar CSV para estado inicial
  console.log('🔄 Resetando CSV para estado inicial...');
  CSVSimulator.resetarCSV();
  
  // 2. Mostrar estado inicial
  console.log('\n📊 ESTADO INICIAL:');
  CSVSimulator.mostrarEstado();
  
  // 3. Criar workbook e simular processamento
  console.log('\n🔄 Simulando processamento com API...');
  const workbook = new MockWorkbook().loadRealData();
  
  // Simular criação de documentos
  atualizarPlanilhaComResultado(workbook, 3, true, '00A0000555');
  atualizarPlanilhaComResultado(workbook, 4, true, '00A0000666');
  
  // 4. Mostrar estado após processamento
  console.log('\n📊 ESTADO APÓS PROCESSAMENTO:');
  CSVSimulator.mostrarEstado();
  
  // 5. Simular segunda execução (deve carregar do CSV)
  console.log('\n🔄 SIMULANDO SEGUNDA EXECUÇÃO...');
  console.log('Carregando dados do CSV...');
  
  const dadosCSV = CSVSimulator.carregarCSV();
  console.log('\nVerificando quais linhas seriam processadas:');
  
  dadosCSV.slice(2).forEach((linha, index) => {
    const numeroLinha = index + 3;
    const notaCriada = linha[21];
    const jaProcessada = notaCriada === 'Sim';
    
    console.log(`  Linha ${numeroLinha}: ${jaProcessada ? '❌ IGNORAR (já processada)' : '✅ PROCESSAR (nova)'}`);
  });
  
  // 6. Processar linha restante
  console.log('\n🔄 Processando linha 5 (não processada)...');
  atualizarPlanilhaComResultado(workbook, 5, true, '00A0000777');
  
  // 7. Estado final
  console.log('\n📊 ESTADO FINAL:');
  CSVSimulator.mostrarEstado();
  
  console.log(`\n✅ Arquivo CSV salvo em: ${CSVSimulator.getCaminhoCSV()}`);
  console.log('🎯 Simulação completa! O CSV agora replica perfeitamente a planilha Excel!');
}

testeCSVSimulator();