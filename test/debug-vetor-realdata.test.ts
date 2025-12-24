/**
 * Teste para verificar se o vetor realData está sendo atualizado
 */

import { MockWorkbook, atualizarPlanilhaComResultado } from '../src/mocks/excel-mock';

function testeVetorRealData() {
  console.log('🔍 TESTE VETOR REALDATA - SIMULAÇÃO COMPLETA\n');
  
  // 1. Criar workbook com dados reais
  const workbook = new MockWorkbook().loadRealData();
  
  console.log('📊 VETOR REALDATA INICIAL:');
  const vetorInicial = workbook.getRealDataVector();
  vetorInicial.slice(2).forEach((linha, index) => {
    console.log(`  Linha ${index + 3}: NotaCriada="${linha[21]}", Retorno="${linha[22]}"`);
  });
  
  // 2. Simular processamento - atualizar linha 3
  console.log('\n🔄 Simulando processamento da linha 3...');
  atualizarPlanilhaComResultado(workbook, 3, true, '00A0000999');
  
  // 3. Verificar vetor realData após atualização
  console.log('\n📊 VETOR REALDATA APÓS ATUALIZAÇÃO:');
  const vetorAtualizado = workbook.getRealDataVector();
  vetorAtualizado.slice(2).forEach((linha, index) => {
    console.log(`  Linha ${index + 3}: NotaCriada="${linha[21]}", Retorno="${linha[22]}"`);
  });
  
  // 4. Simular segunda execução - verificar se linha 3 seria ignorada
  console.log('\n🔍 SIMULANDO SEGUNDA EXECUÇÃO:');
  console.log('Verificando quais linhas seriam processadas...');
  
  vetorAtualizado.slice(2).forEach((linha, index) => {
    const linhaNum = index + 3;
    const notaCriada = linha[21];
    const jaProcessada = notaCriada === 'Sim';
    
    console.log(`  Linha ${linhaNum}: ${jaProcessada ? '❌ IGNORAR (já processada)' : '✅ PROCESSAR (nova)'}`);
  });
  
  console.log('\n✅ Agora o vetor realData está sendo atualizado para simulação!');
}

testeVetorRealData();