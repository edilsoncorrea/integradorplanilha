/**
 * Simulador interativo para criação de pedidos
 */

import { MockWorkbook } from '../mocks/excel-mock';
import { main } from '../office-scripts/PedidoDeVenda';
import { apiDebugger } from '../debug/api-debugger';

export class PedidoSimulator {
  private workbook: MockWorkbook;

  constructor() {
    this.workbook = new MockWorkbook();
    this.workbook.loadRealData();
  }

  // Gera pedidos da planilha
  gerarPedidosDaPlanilha() {
    console.log('🛒 Gerando pedidos da planilha...\n');
    
    const result = main(this.workbook as any, { action: 'buildPedidoVendaFromSheet' });
    
    if (result.payloads && result.payloads.length > 0) {
      console.log(`✅ ${result.payloads.length} pedido(s) gerado(s):\n`);
      
      result.payloads.forEach((pedido: any, index: number) => {
        console.log(`📋 Pedido ${index + 1} (Linha ${pedido.sheetRow}):`);
        console.log(`   Empresa: ${pedido.payload.CodigoEmpresa}`);
        console.log(`   Cliente: ${pedido.payload.IdentificadorCliente}`);
        console.log(`   Operação: ${pedido.payload.IdentificadorOperacao}`);
        console.log(`   Data: ${pedido.payload.DataEmissao}`);
        console.log(`   Valor Total: ${pedido.payload.Itens[0].Valor}`);
        console.log(`   Status: ${pedido.payload.Status}`);
        console.log('');
      });
      
      return result.payloads;
    } else {
      console.log('❌ Nenhum pedido gerado');
      return [];
    }
  }

  // Simula envio para API
  simularEnvioAPI(pedidos: any[]) {
    console.log('🚀 Simulando envio para API...\n');
    
    const resultados = pedidos.map((pedido: any, index: number) => {
      const endpoint = 'https://087344bimerapi.alterdata.cloud/api/pedidos';
      
      // 🔴 BREAKPOINT: Antes da chamada da API
      apiDebugger.beforeAPICall(endpoint, pedido.payload, `Pedido ${index + 1} - Linha ${pedido.sheetRow}`);
      
      // Simula resposta da API
      const sucesso = Math.random() > 0.2; // 80% de sucesso
      
      const response = {
        success: sucesso,
        pedidoId: sucesso ? `PV${Date.now()}${pedido.sheetRow}` : null,
        message: sucesso ? 'Pedido criado com sucesso' : 'Erro: Cliente inválido',
        timestamp: new Date().toISOString()
      };
      
      if (sucesso) {
        // 🟢 BREAKPOINT: Após resposta da API (sucesso)
        apiDebugger.afterAPIResponse(endpoint, response, `Pedido ${index + 1} - Sucesso`);
      } else {
        // 🔴 BREAKPOINT: Erro na API
        apiDebugger.onAPIError(endpoint, response, `Pedido ${index + 1} - Erro`);
      }
      
      return {
        sheetRow: pedido.sheetRow,
        pedidoCriado: response.pedidoId || '',
        retorno: response.message
      };
    });
    
    console.log('📊 Resultados do envio:');
    resultados.forEach((resultado: any) => {
      console.log(`   Linha ${resultado.sheetRow}: ${resultado.retorno}`);
      if (resultado.pedidoCriado) {
        console.log(`   ID Pedido: ${resultado.pedidoCriado}`);
      }
    });
    console.log('');
    
    return resultados;
  }

  // Executa fluxo completo
  executarFluxoCompleto() {
    console.log('🎯 Executando fluxo completo de criação de pedidos...\n');
    
    // 1. Gerar pedidos
    const pedidos = this.gerarPedidosDaPlanilha();
    
    if (pedidos.length === 0) {
      console.log('❌ Nenhum pedido para processar');
      return;
    }
    
    // 2. Simular envio
    const resultados = this.simularEnvioAPI(pedidos);
    
    // 3. Mostrar resumo
    const sucessos = resultados.filter((r: any) => r.pedidoCriado).length;
    const erros = resultados.length - sucessos;
    
    console.log('📈 Resumo Final:');
    console.log(`   ✅ Sucessos: ${sucessos}`);
    console.log(`   ❌ Erros: ${erros}`);
    console.log(`   📊 Taxa de sucesso: ${((sucessos / resultados.length) * 100).toFixed(1)}%`);
    
    return { pedidos, resultados, sucessos, erros };
  }

  // Obter exemplo de pedido
  obterExemploPedido() {
    console.log('📝 Obtendo exemplo de pedido...\n');
    
    const result = main(this.workbook as any, { action: 'getSamplePedido' });
    
    if (result.sample) {
      console.log('✅ Exemplo de pedido:');
      console.log(JSON.stringify(result.sample, null, 2));
    }
    
    return result.sample;
  }
}