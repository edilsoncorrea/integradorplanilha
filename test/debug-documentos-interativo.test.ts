/**
 * Debug INTERATIVO - Criar Documentos com Breakpoints
 * Use Ctrl+Shift+D no VS Code e selecione "📄 Debug Documentos Interativo"
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { criarDocumento } from '../src/office-scripts/DocumentoCompleto';
import { RealAPIClient } from '../src/api/real-api-client';
import { apiDebugger } from '../src/debug/api-debugger';

async function debugDocumentosInterativo() {
  console.log('📄 DEBUG INTERATIVO - CRIAR DOCUMENTOS');
  console.log('=' .repeat(60));
  
  // 🔴 BREAKPOINT 1: Carregamento de dados
  debugger; // <- Pare aqui para inspecionar
  console.log('\n📊 Carregando dados da planilha CSV...');
  
  const mockWorkbook = new MockWorkbook().loadRealData();
  const sheet = mockWorkbook.getWorksheet('Documento');
  const data = sheet?.getUsedRange()?.getValues();
  
  console.log(`✅ ${data?.length || 0} linhas carregadas`);
  
  // 🔴 BREAKPOINT 2: Geração de payloads de documentos
  debugger; // <- Pare aqui para ver dados carregados
  console.log('\n📄 Gerando payloads dos documentos...');
  
  const result = criarDocumento(mockWorkbook);
  
  console.log(`✅ ${result.payloads?.length || 0} documentos preparados`);
  
  if (result.payloads?.length > 0) {
    console.log('\n📦 Primeiro documento:');
    console.log(`Endpoint: ${result.payloads[0].endpoint}`);
    console.log(`Linha: ${result.payloads[0].sheetRow}`);
    
    // Mostrar diferenças vs pedidos
    const payload = result.payloads[0].payload;
    console.log('\n🔍 Campos específicos de documentos:');
    console.log(`StatusNotaFiscalEletronica: ${payload.StatusNotaFiscalEletronica}`);
    console.log(`TipoDocumento: ${payload.TipoDocumento}`);
    console.log(`TipoPagamento: ${payload.TipoPagamento}`);
    console.log(`AliquotaConvenio: ${payload.Pagamentos[0]?.AliquotaConvenio}`);
  }
  
  // 🔴 BREAKPOINT 3: Autenticação
  debugger; // <- Pare aqui para inspecionar payloads de documentos
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
  
  // 🔴 BREAKPOINT 4: Envio de documentos
  debugger; // <- Pare aqui antes de enviar documentos
  console.log('\n📄 Enviando documentos para API...');
  
  for (let i = 0; i < Math.min(result.payloads.length, 2); i++) {
    const documento = result.payloads[i];
    
    console.log(`\n📤 Documento ${i + 1} (linha ${documento.sheetRow})`);
    console.log(`   Endpoint: ${documento.endpoint}`);
    console.log(`   Empresa: ${documento.payload.CodigoEmpresa}`);
    console.log(`   Cliente: ${documento.payload.IdentificadorPessoa}`);
    console.log(`   Tipo: ${documento.payload.TipoDocumento}`);
    
    // 🔴 BREAKPOINT 5: Antes de cada documento
    debugger; // <- Pare aqui para inspecionar cada documento
    
    try {
      // Para testar com API real, descomente a linha abaixo:
      const documentoResult = await apiClient.callAPI('/Documento', 'POST', documento.payload);
      
      if (documentoResult.success && documentoResult.data?.ListaObjetos?.[0]?.Identificador) {
        const identificador = documentoResult.data.ListaObjetos[0].Identificador;
        console.log(`   ✅ Documento criado: ${identificador}`);
        
        // IMPORTANTE: Atualizar CSV com resultado
        const { atualizarPlanilhaComResultado } = require('../src/mocks/excel-mock');
        atualizarPlanilhaComResultado(mockWorkbook, documento.sheetRow, true, identificador);
        
      } else {
        console.log(`   ❌ Erro: ${JSON.stringify(documentoResult.error)}`);
        
        // Atualizar CSV com erro
        const { atualizarPlanilhaComResultado } = require('../src/mocks/excel-mock');
        atualizarPlanilhaComResultado(mockWorkbook, documento.sheetRow, false, 'ERRO');
      }
      
    } catch (error) {
      console.log(`   ❌ Exceção: ${error}`);
      
      // Atualizar CSV com erro
      const { atualizarPlanilhaComResultado } = require('../src/mocks/excel-mock');
      atualizarPlanilhaComResultado(mockWorkbook, documento.sheetRow, false, 'ERRO');
    }
  }
  
  // 🔴 BREAKPOINT 6: Final
  debugger; // <- Pare aqui para ver resultados finais
  console.log('\n🎯 Debug interativo de documentos concluído!');
  console.log(`Total processado: ${result.payloads.length} documentos`);
  
  console.log('\n📋 Resumo das diferenças:');
  console.log('   📄 Documentos → /api/documentos (notas fiscais)');
  console.log('   📦 Pedidos → /api/venda/pedidos (pedidos de venda)');
  console.log('   🔧 Campo extra: AliquotaConvenio nos documentos');
  console.log('   📊 StatusNotaFiscalEletronica: "A" nos documentos');
}

// Executar debug interativo
debugDocumentosInterativo().catch(console.error);