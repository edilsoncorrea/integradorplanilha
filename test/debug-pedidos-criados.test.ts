/**
 * Debug para verificar identificadores e códigos dos pedidos criados
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { main as pedidoMain } from '../src/office-scripts/PedidoDeVenda';
import { RealAPIClient } from '../src/api/real-api-client';

async function debugPedidosCriados() {
  console.log('🔍 DEBUG - VERIFICAR PEDIDOS CRIADOS\n');
  console.log('=' .repeat(60));
  
  // PASSO 1: Autenticar
  console.log('\n🔐 Autenticando...');
  const apiClient = new RealAPIClient();
  const authResult = await apiClient.authenticate();
  
  if (!authResult.success) {
    console.log(`❌ Falha na autenticação: ${authResult.error}`);
    return;
  }
  
  console.log('✅ Autenticado com sucesso!');
  
  // PASSO 2: Gerar um pedido de teste
  console.log('\n📦 Gerando pedido de teste...');
  const mockWorkbook = new MockWorkbook().loadRealData();
  const result = pedidoMain(mockWorkbook, { action: 'buildPedidoVendaFromSheet' });
  
  if (!result.payloads || result.payloads.length === 0) {
    console.log('❌ Nenhum payload gerado');
    return;
  }
  
  const pedidoTeste = result.payloads[0];
  console.log(`✅ Payload gerado para linha ${pedidoTeste.sheetRow}`);
  
  // PASSO 3: Criar pedido e capturar resposta detalhada
  console.log('\n🚀 Criando pedido na API...');
  console.log(`Empresa: ${pedidoTeste.payload.CodigoEmpresa}`);
  console.log(`Cliente: ${pedidoTeste.payload.IdentificadorCliente}`);
  
  try {
    const pedidoResult = await apiClient.createPedido(pedidoTeste.payload);
    
    console.log('\n📊 RESPOSTA COMPLETA DA API:');
    console.log(JSON.stringify(pedidoResult, null, 2));
    
    // PASSO 4: Extrair identificadores
    if (pedidoResult.success && pedidoResult.data) {
      console.log('\n✅ PEDIDO CRIADO COM SUCESSO!');
      
      const responseData = pedidoResult.data;
      
      // Verificar estrutura da resposta
      console.log('\n🔍 ANÁLISE DA RESPOSTA:');
      console.log(`Tem Erros: ${responseData.Erros ? responseData.Erros.length : 'N/A'}`);
      console.log(`Tem ListaObjetos: ${responseData.ListaObjetos ? responseData.ListaObjetos.length : 'N/A'}`);
      
      if (responseData.ListaObjetos && responseData.ListaObjetos.length > 0) {
        const pedidoCriado = responseData.ListaObjetos[0];
        
        console.log('\n🎯 DADOS DO PEDIDO CRIADO:');
        console.log(`📋 Identificador: ${pedidoCriado.Identificador || 'N/A'}`);
        console.log(`🔢 Código: ${pedidoCriado.Codigo || 'N/A'}`);
        
        // Verificar outros campos possíveis
        console.log('\n📝 TODOS OS CAMPOS RETORNADOS:');
        Object.keys(pedidoCriado).forEach(key => {
          console.log(`   ${key}: ${pedidoCriado[key]}`);
        });
        
        // PASSO 5: Testar consulta do pedido criado
        console.log('\n🔍 TESTANDO CONSULTA DO PEDIDO CRIADO...');
        
        if (pedidoCriado.Identificador) {
          try {
            // Tentar consultar o pedido pelo identificador
            const consultaResult = await apiClient.callAPI(
              `/api/venda/pedidos/${pedidoCriado.Identificador}`, 
              {}, 
              'Consultar Pedido Criado'
            );
            
            console.log('\n📋 CONSULTA DO PEDIDO:');
            console.log(JSON.stringify(consultaResult, null, 2));
            
          } catch (error) {
            console.log(`❌ Erro na consulta: ${error}`);
          }
        }
        
      } else {
        console.log('❌ Nenhum objeto retornado na ListaObjetos');
      }
      
    } else {
      console.log('\n❌ FALHA NA CRIAÇÃO DO PEDIDO:');
      console.log(`Erro: ${JSON.stringify(pedidoResult.error, null, 2)}`);
    }
    
  } catch (error) {
    console.log(`\n❌ EXCEÇÃO: ${error}`);
  }
  
  console.log('\n' + '=' .repeat(60));
  console.log('🎯 Verificação concluída!');
}

debugPedidosCriados().catch(console.error);