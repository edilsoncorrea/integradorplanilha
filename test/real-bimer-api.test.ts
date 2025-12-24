/**
 * Teste com API Bimer REAL usando dados CSV
 */

import { RealAPIClient } from '../src/api/real-api-client';
import { PedidoSimulator } from '../src/simulators/pedido-simulator';

async function testRealBimerAPI() {
  console.log('🌐 TESTE COM API BIMER REAL\n');
  console.log('=' .repeat(60));
  
  const apiClient = new RealAPIClient();
  
  try {
    // 1. AUTENTICAÇÃO REAL
    console.log('🔐 PASSO 1: Autenticação com API Bimer...');
    const authResult = await apiClient.authenticate();
    
    if (!authResult.success) {
      console.log('❌ Falha na autenticação:', authResult.error);
      return;
    }
    
    console.log('✅ Autenticação bem-sucedida!');
    console.log(`Token: ${authResult.token?.substring(0, 30)}...`);
    
    // 2. GERAR PEDIDOS DOS DADOS CSV
    console.log('\n📋 PASSO 2: Gerando pedidos dos dados CSV...');
    const simulator = new PedidoSimulator();
    const pedidos = simulator.gerarPedidosDaPlanilha();
    
    if (pedidos.length === 0) {
      console.log('❌ Nenhum pedido gerado dos dados CSV');
      return;
    }
    
    console.log(`✅ ${pedidos.length} pedido(s) gerado(s) dos dados CSV`);
    
    // 3. ENVIAR PEDIDOS PARA API REAL
    console.log('\n🚀 PASSO 3: Enviando pedidos para API Bimer...');
    
    const resultados = [];
    
    for (let i = 0; i < pedidos.length; i++) {
      const pedido = pedidos[i];
      console.log(`\n--- ENVIANDO PEDIDO ${i + 1}/${pedidos.length} ---`);
      console.log(`Linha da planilha: ${pedido.sheetRow}`);
      console.log(`Cliente: ${pedido.payload.IdentificadorCliente}`);
      console.log(`Valor: ${pedido.payload.Itens[0].Valor}`);
      
      try {
        const result = await apiClient.callAPI(
          '/api/pedidos', 
          pedido.payload, 
          `Pedido ${i + 1} - Linha ${pedido.sheetRow}`
        );
        
        if (result.success) {
          console.log(`✅ Sucesso! ID: ${(result.data as any)?.id || 'N/A'}`);
          resultados.push({
            linha: pedido.sheetRow,
            status: 'sucesso',
            id: (result.data as any)?.id,
            response: result.data
          });
        } else {
          console.log(`❌ Erro: ${JSON.stringify(result.error)}`);
          resultados.push({
            linha: pedido.sheetRow,
            status: 'erro',
            error: result.error
          });
        }
        
      } catch (error) {
        console.log(`💥 Exceção: ${error}`);
        resultados.push({
          linha: pedido.sheetRow,
          status: 'exceção',
          error: String(error)
        });
      }
      
      // Pausa entre chamadas para não sobrecarregar API
      await new Promise(resolve => setTimeout(resolve, 1000));
    }
    
    // 4. RESUMO FINAL
    console.log('\n📊 PASSO 4: Resumo final...');
    console.log('=' .repeat(60));
    
    const sucessos = resultados.filter(r => r.status === 'sucesso').length;
    const erros = resultados.filter(r => r.status === 'erro').length;
    const excecoes = resultados.filter(r => r.status === 'exceção').length;
    
    console.log(`📋 Total processado: ${resultados.length}`);
    console.log(`✅ Sucessos: ${sucessos}`);
    console.log(`❌ Erros: ${erros}`);
    console.log(`💥 Exceções: ${excecoes}`);
    console.log(`📈 Taxa de sucesso: ${((sucessos / resultados.length) * 100).toFixed(1)}%`);
    
    // Mostrar detalhes dos erros
    if (erros > 0 || excecoes > 0) {
      console.log('\n🔍 Detalhes dos erros:');
      resultados.forEach(r => {
        if (r.status !== 'sucesso') {
          console.log(`   Linha ${r.linha}: ${r.status} - ${JSON.stringify(r.error).substring(0, 100)}...`);
        }
      });
    }
    
  } catch (error) {
    console.log('\n💥 Erro geral no teste:', error);
  }
  
  console.log('\n' + '=' .repeat(60));
  console.log('🌐 Teste com API Bimer real concluído!');
}

testRealBimerAPI().catch(console.error);