/**
 * Testes para criação de pedidos de venda
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { main } from '../src/office-scripts/PedidoDeVenda';

(global as any).ExcelScript = {};

function runPedidoVendaTests() {
  console.log('🛒 Testando criação de pedidos de venda...\n');

  // Teste 1: buildPedidoVendaFromSheet
  console.log('📋 Teste 1: buildPedidoVendaFromSheet');
  const mockWorkbook = new MockWorkbook();
  mockWorkbook.loadRealData();
  
  const result = main(mockWorkbook as any, { action: 'buildPedidoVendaFromSheet' });
  
  console.log(`✅ Pedidos gerados: ${result.payloads?.length || 0}`);
  
  if (result.payloads && result.payloads.length > 0) {
    console.log('📄 Primeiro pedido:');
    const pedido = result.payloads[0];
    console.log(`   Linha: ${pedido.sheetRow}`);
    console.log(`   Empresa: ${pedido.payload.CodigoEmpresa}`);
    console.log(`   Cliente: ${pedido.payload.IdentificadorCliente}`);
    console.log(`   Data: ${pedido.payload.DataEmissao}`);
    console.log(`   Itens: ${pedido.payload.Itens.length}`);
    console.log(`   Valor: ${pedido.payload.Itens[0].Valor}`);
  }
  console.log('');

  // Teste 2: getSamplePedido
  console.log('📝 Teste 2: getSamplePedido');
  const sampleResult = main(mockWorkbook as any, { action: 'getSamplePedido' });
  
  if (sampleResult.sample) {
    console.log('✅ Sample pedido carregado');
    console.log(`   Empresa: ${sampleResult.sample.CodigoEmpresa}`);
    console.log(`   Cliente: ${sampleResult.sample.IdentificadorCliente}`);
    console.log(`   Itens: ${sampleResult.sample.Itens.length}`);
    console.log(`   Pagamentos: ${sampleResult.sample.Pagamentos.length}`);
  }
  console.log('');

  // Teste 3: Validação estrutura pedido
  console.log('🔍 Teste 3: Validação estrutura');
  if (result.payloads && result.payloads.length > 0) {
    const pedido = result.payloads[0].payload;
    const requiredFields = [
      'CodigoEmpresa',
      'DataEmissao',
      'IdentificadorOperacao',
      'IdentificadorCliente',
      'Itens',
      'Status',
      'TipoFrete'
    ];
    
    const missing = requiredFields.filter(field => !(field in pedido));
    console.log(`${missing.length === 0 ? '✅' : '❌'} Campos obrigatórios: ${missing.length === 0 ? 'OK' : 'Faltando: ' + missing.join(', ')}`);
    
    if (pedido.Itens && pedido.Itens.length > 0) {
      const item = pedido.Itens[0];
      const itemFields = ['IdentificadorProduto', 'QuantidadePedida', 'Valor', 'ValorUnitario'];
      const missingItem = itemFields.filter(field => !(field in item));
      console.log(`${missingItem.length === 0 ? '✅' : '❌'} Estrutura itens: ${missingItem.length === 0 ? 'OK' : 'Faltando: ' + missingItem.join(', ')}`);
    }
  }
  console.log('');

  console.log('🎉 Testes de pedido de venda concluídos!');
}

runPedidoVendaTests();