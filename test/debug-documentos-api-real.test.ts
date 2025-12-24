/**
 * Debug para testar chamada REAL da API Bimer - Criar Documentos
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { criarDocumento } from '../src/office-scripts/DocumentoCompleto';
import { RealAPIClient } from '../src/api/real-api-client';
import { apiDebugger } from '../src/debug/api-debugger';

async function debugDocumentosAPIReal() {
  console.log('📄 DEBUG - API REAL CRIAR DOCUMENTOS\n');
  console.log('=' .repeat(70));
  
  // PASSO 1: Autenticar na API
  console.log('\n🔐 PASSO 1: Autenticando na API Bimer');
  const apiClient = new RealAPIClient();
  
  // Ativar breakpoints para monitorar chamadas
  apiDebugger.enableBreakpoint('AUTH_CHECK');
  apiDebugger.enableBreakpoint('BEFORE_API_CALL');
  apiDebugger.enableBreakpoint('AFTER_API_RESPONSE');
  
  const authResult = await apiClient.authenticate();
  
  if (!authResult.success) {
    console.log(`❌ Falha na autenticação: ${authResult.error}`);
    return;
  }
  
  console.log(`✅ Autenticado com sucesso!`);
  console.log(`🎫 Token: ${authResult.token?.substring(0, 20)}...`);
  
  // PASSO 2: Gerar payload do documento
  console.log('\n📄 PASSO 2: Gerando payload do documento');
  const mockWorkbook = new MockWorkbook().loadRealData();
  const result = criarDocumento(mockWorkbook);
  
  if (!result.payloads || result.payloads.length === 0) {
    console.log('❌ Nenhum documento para processar');
    return;
  }
  
  console.log(`✅ ${result.payloads.length} documentos preparados`);
  
  // PASSO 3: Processar TODOS os documentos
  console.log('\n🚀 PASSO 3: Processando TODOS os documentos');
  
  for (let i = 0; i < result.payloads.length; i++) {
    const documento = result.payloads[i];
    
    console.log(`\n📄 DOCUMENTO ${i + 1}/${result.payloads.length} (linha ${documento.sheetRow})`);
    console.log(`📍 Endpoint: ${documento.endpoint}`);
    console.log(`Empresa: ${documento.payload.CodigoEmpresa}`);
    console.log(`Cliente: ${documento.payload.IdentificadorPessoa}`);
    console.log(`Tipo: ${documento.payload.TipoDocumento}`);
    
    try {
      const documentoResult = await apiClient.callAPI(
        documento.endpoint, 
        documento.payload, 
        `Criar Documento ${i + 1}`
      );
    
      // Analisar resultado
      if (documentoResult.success && documentoResult.data) {
        console.log(`   ✅ SUCESSO!`);
        
        const responseData = documentoResult.data;
        
        if (responseData.ListaObjetos && responseData.ListaObjetos.length > 0) {
          const documentoCriado = responseData.ListaObjetos[0];
          const identificador = documentoCriado.Identificador;
          
          console.log(`   📋 Identificador: ${identificador}`);
          
          // IMPORTANTE: Atualizar CSV com resultado da API
          if (identificador) {
            const { atualizarPlanilhaComResultado } = require('../src/mocks/excel-mock');
            atualizarPlanilhaComResultado(mockWorkbook, documento.sheetRow, true, identificador);
            console.log(`   💾 CSV atualizado: NotaCriada="Sim", Retorno="${identificador}"`);
          }
          
        } else {
          console.log('   ❌ Nenhum objeto retornado');
        }
        
      } else {
        console.log(`   ❌ FALHA: ${JSON.stringify(documentoResult.error)}`);
        
        // IMPORTANTE: Atualizar CSV com erro
        const { atualizarPlanilhaComResultado } = require('../src/mocks/excel-mock');
        atualizarPlanilhaComResultado(mockWorkbook, documento.sheetRow, false, 'ERRO');
        console.log(`   💾 CSV atualizado: NotaCriada="Não", Retorno="ERRO"`);
      }
      
    } catch (error) {
      console.log(`   ❌ EXCEÇÃO: ${error}`);
      
      // IMPORTANTE: Atualizar CSV com erro de exceção
      const { atualizarPlanilhaComResultado } = require('../src/mocks/excel-mock');
      atualizarPlanilhaComResultado(mockWorkbook, documento.sheetRow, false, 'ERRO');
      console.log(`   💾 CSV atualizado: NotaCriada="Não", Retorno="ERRO"`);
    }
  }
  
  console.log('\n' + '=' .repeat(70));
  console.log('🎯 Debug API Real - Documentos concluído!');
  
  console.log('\n📋 Comparação com Pedidos:');
  console.log('   📄 Documentos: /api/documentos → Notas Fiscais');
  console.log('   📦 Pedidos: /api/venda/pedidos → Pedidos de Venda');
  console.log('   🔧 Campos extras em documentos:');
  console.log('      - StatusNotaFiscalEletronica');
  console.log('      - TipoDocumento');
  console.log('      - TipoPagamento');
  console.log('      - AliquotaConvenio (nos pagamentos)');
}

debugDocumentosAPIReal().catch(console.error);