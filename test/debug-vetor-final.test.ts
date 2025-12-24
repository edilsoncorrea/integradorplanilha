/**
 * Debug: Verificar estado final do vetor após processamento
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { main as documentoMain } from '../src/office-scripts/DocumentoCompleto';
import { RealAPIClient } from '../src/api/real-api-client';
import { atualizarPlanilhaComResultado } from '../src/mocks/excel-mock';

async function debugVetorFinal() {
  console.log('🔍 DEBUG VETOR FINAL\n');
  
  // 1. Carregar dados iniciais
  const workbook = new MockWorkbook().loadRealData();
  const sheet = workbook.getWorksheet('Documento');
  
  console.log('📊 VETOR INICIAL:');
  const dadosIniciais = sheet?.getUsedRange()?.getValues();
  if (dadosIniciais) {
    dadosIniciais.slice(2).forEach((linha, index) => {
      console.log(`  Linha ${index + 3}: NotaCriada="${linha[21]}", Retorno="${linha[22]}"`);
    });
  }
  
  // 2. Processar documentos
  console.log('\n📄 Processando documentos...');
  const result = documentoMain(workbook, { action: 'criarDocumento' });
  
  const apiClient = new RealAPIClient();
  await apiClient.authenticate();
  
  for (const doc of result.payloads) {
    try {
      const response = await apiClient.callAPI('/Documento', 'POST', doc.payload);
      const identificador = response.data.ListaObjetos[0].Identificador;
      
      // Atualizar planilha
      atualizarPlanilhaComResultado(workbook, doc.sheetRow, true, identificador);
      
    } catch (error) {
      atualizarPlanilhaComResultado(workbook, doc.sheetRow, false, 'ERRO');
    }
  }
  
  // 3. Mostrar vetor final
  console.log('\n📊 VETOR FINAL ATUALIZADO:');
  const dadosFinais = sheet?.getUsedRange()?.getValues();
  if (dadosFinais) {
    dadosFinais.slice(2).forEach((linha, index) => {
      console.log(`  Linha ${index + 3}: NotaCriada="${linha[21]}", Retorno="${linha[22]}"`);
    });
  }
  
  console.log('\n✅ O vetor foi atualizado com as informações da API!');
}

debugVetorFinal().catch(console.error);