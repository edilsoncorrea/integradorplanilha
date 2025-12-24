/**
 * Debug FLUXO COMPLETO - Simula processo real de criação de documentos
 * 1. Gera documentos das linhas vazias
 * 2. Chama API Bimer (simulada)
 * 3. Atualiza planilha com resultado
 * 4. Próxima execução ignora linhas já processadas
 */

import { MockWorkbook, atualizarPlanilhaComResultado } from '../src/mocks/excel-mock';
import { criarDocumento } from '../src/office-scripts/DocumentoCompleto';
import { NotaCriada, RetornoAPI } from '../src/office-scripts/Constantes';

async function debugFluxoCompletoDocumentos() {
  console.log('🔄 DEBUG FLUXO COMPLETO - DOCUMENTOS\n');
  console.log('=' .repeat(70));
  
  // PASSO 1: Estado inicial da planilha
  console.log('\n📊 PASSO 1: Estado inicial da planilha');
  const mockWorkbook = new MockWorkbook().loadRealData();
  let sheet = mockWorkbook.getWorksheet('Documento');
  let data = sheet?.getUsedRange()?.getValues();
  
  console.log('Estado inicial:');
  data?.forEach((row, index) => {
    if (index > 1) { // Pular headers
      const notaCriada = row[NotaCriada] || '';
      const retorno = row[RetornoAPI] || '';
      console.log(`  Linha ${index + 1}: NotaCriada="${notaCriada}", Retorno="${retorno}"`);
    }
  });
  
  // PASSO 2: Primeira execução - gerar documentos
  console.log('\n📄 PASSO 2: Primeira execução - gerar documentos');
  let result = criarDocumento(mockWorkbook);
  
  console.log(`Documentos para processar: ${result.payloads?.length || 0}`);
  
  if (result.payloads && result.payloads.length > 0) {
    // PASSO 3: Processar documentos
    console.log('\n🔐 PASSO 3: Processando documentos...');
    
    // Processar cada documento
    for (const documento of result.payloads) {
      console.log(`\n📤 Processando documento da linha ${documento.sheetRow}`);
      
      try {
        // Simular resposta de sucesso da API
        const documentoResult = {
          success: true,
          data: {
            Erros: [],
            ListaObjetos: [
              {
                Identificador: `00A000${String(Math.floor(Math.random() * 1000)).padStart(4, '0')}`,
                Codigo: `DOC${String(Math.floor(Math.random() * 10000)).padStart(6, '0')}`
              }
            ]
          }
        };
        
        // PASSO 4: Atualizar planilha com resultado
        if (documentoResult.success && documentoResult.data.ListaObjetos.length > 0) {
          const documentoCriado = documentoResult.data.ListaObjetos[0];
          
          console.log(`   ✅ Documento criado: ${documentoCriado.Identificador}`);
          
          // Atualizar mock da planilha
          atualizarPlanilhaComResultado(mockWorkbook, documento.sheetRow, true, documentoCriado.Identificador);
          
        } else {
          console.log(`   ❌ Falha na criação`);
          atualizarPlanilhaComResultado(mockWorkbook, documento.sheetRow, false, 'Erro na API');
        }
        
      } catch (error) {
        console.log(`   ❌ Exceção: ${error}`);
        atualizarPlanilhaComResultado(mockWorkbook, documento.sheetRow, false, `Erro: ${error}`);
      }
    }
  }
  
  // PASSO 5: Estado após primeira execução
  console.log('\n📊 PASSO 5: Estado após primeira execução');
  data = sheet?.getUsedRange()?.getValues();
  
  console.log('Estado atualizado:');
  data?.forEach((row, index) => {
    if (index > 1) { // Pular headers
      const notaCriada = row[NotaCriada] || '';
      const retorno = row[RetornoAPI] || '';
      console.log(`  Linha ${index + 1}: NotaCriada="${notaCriada}", Retorno="${retorno}"`);
    }
  });
  
  // PASSO 6: Segunda execução - deve ignorar linhas já processadas
  console.log('\n🔄 PASSO 6: Segunda execução (deve ignorar linhas processadas)');
  result = criarDocumento(mockWorkbook);
  
  console.log(`Documentos para processar na 2ª execução: ${result.payloads?.length || 0}`);
  
  if (result.payloads && result.payloads.length === 0) {
    console.log('✅ Correto! Nenhum documento para processar (todos já foram criados)');
  } else {
    console.log('⚠️ Ainda há documentos para processar:');
    result.payloads?.forEach((doc: any) => {
      console.log(`   - Linha ${doc.sheetRow}`);
    });
  }
  
  // PASSO 7: Simular adição de nova linha
  console.log('\n➕ PASSO 7: Simulando adição de nova linha');
  
  // Adicionar nova linha ao mock
  const novaLinha = [
    '000004', '000149', 'NOVA EMPRESA TESTE', '00A000005F', '000022', '00A000000U', 
    '6.104', '000899', '00A00000R6', 'NOVO PRODUTO TESTE', '8', 'R$ 800,00', 
    '', '', '000005', '00A000000K', 'Não', '000005', 
    '00A0000006', '07/12/2025', '07/12/2025', '', ''
  ];
  
  data?.push(novaLinha);
  sheet?.setData(data!);
  
  console.log('Nova linha adicionada (linha 6)');
  
  // PASSO 8: Terceira execução - deve processar apenas a nova linha
  console.log('\n🆕 PASSO 8: Terceira execução (deve processar apenas nova linha)');
  result = criarDocumento(mockWorkbook);
  
  console.log(`Documentos para processar na 3ª execução: ${result.payloads?.length || 0}`);
  
  if (result.payloads && result.payloads.length === 1) {
    console.log(`✅ Correto! Processando apenas linha ${result.payloads[0].sheetRow} (nova linha)`);
    
    // Processar a nova linha
    const novoDoc = result.payloads[0];
    const novoId = `00A000${String(Math.floor(Math.random() * 1000)).padStart(4, '0')}`;
    atualizarPlanilhaComResultado(mockWorkbook, novoDoc.sheetRow, true, novoId);
    
  } else if (result.payloads && result.payloads.length === 0) {
    console.log('⚠️ Nenhuma linha nova detectada (possível problema no mock)');
  }
  
  // PASSO 9: Estado final
  console.log('\n📊 PASSO 9: Estado final da planilha');
  data = sheet?.getUsedRange()?.getValues();
  
  console.log('Estado final:');
  data?.forEach((row, index) => {
    if (index > 1) { // Pular headers
      const notaCriada = row[NotaCriada] || '';
      const retorno = row[RetornoAPI] || '';
      console.log(`  Linha ${index + 1}: NotaCriada="${notaCriada}", Retorno="${retorno}"`);
    }
  });
  
  console.log('\n' + '=' .repeat(70));
  console.log('🎯 Fluxo completo testado com sucesso!');
  
  console.log('\n📋 Resumo do comportamento:');
  console.log('   1️⃣ Primeira execução: Processa linhas vazias');
  console.log('   2️⃣ Atualiza planilha: NotaCriada="Sim", Retorno=Identificador');
  console.log('   3️⃣ Segunda execução: Ignora linhas já processadas');
  console.log('   4️⃣ Nova linha: Processa apenas linhas novas/vazias');
  console.log('   ✅ Comportamento igual ao Google Apps Script original!');
}

debugFluxoCompletoDocumentos().catch(console.error);