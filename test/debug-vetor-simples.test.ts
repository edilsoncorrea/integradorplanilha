/**
 * Teste simples para verificar se o vetor está sendo atualizado
 */

import { MockWorkbook, atualizarPlanilhaComResultado } from '../src/mocks/excel-mock';

function testeVetorSimples() {
  console.log('🔍 TESTE SIMPLES - ATUALIZAÇÃO DO VETOR\n');
  
  // 1. Criar workbook com dados reais
  const workbook = new MockWorkbook().loadRealData();
  const sheet = workbook.getWorksheet('Documento');
  
  console.log('📊 ANTES da atualização:');
  const dadosAntes = sheet?.getUsedRange()?.getValues();
  if (dadosAntes) {
    console.log(`Linha 3: NotaCriada="${dadosAntes[2][21]}", Retorno="${dadosAntes[2][22]}"`);
    console.log(`Linha 4: NotaCriada="${dadosAntes[3][21]}", Retorno="${dadosAntes[3][22]}"`);
    console.log(`Linha 5: NotaCriada="${dadosAntes[4][21]}", Retorno="${dadosAntes[4][22]}"`);
  }
  
  // 2. Atualizar linha 3
  console.log('\n🔄 Atualizando linha 3...');
  atualizarPlanilhaComResultado(workbook, 3, true, '00A0000123');
  
  // 3. Verificar se foi atualizado
  console.log('\n📊 DEPOIS da atualização:');
  const dadosDepois = sheet?.getUsedRange()?.getValues();
  if (dadosDepois) {
    console.log(`Linha 3: NotaCriada="${dadosDepois[2][21]}", Retorno="${dadosDepois[2][22]}"`);
    console.log(`Linha 4: NotaCriada="${dadosDepois[3][21]}", Retorno="${dadosDepois[3][22]}"`);
    console.log(`Linha 5: NotaCriada="${dadosDepois[4][21]}", Retorno="${dadosDepois[4][22]}"`);
  }
  
  // 4. Verificar se realmente mudou
  const mudou = dadosDepois?.[2][21] === 'Sim' && dadosDepois?.[2][22] === '00A0000123';
  console.log(`\n${mudou ? '✅' : '❌'} Vetor ${mudou ? 'FOI' : 'NÃO FOI'} atualizado!`);
}

testeVetorSimples();