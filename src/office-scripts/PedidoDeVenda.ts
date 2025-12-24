// PedidoDeVenda.ts
// Conversão de funções relacionadas a pedidos de venda (criarPedidoDevenda3, criarPedidoVenda2)
// Actions:
//  - 'buildPedidoVendaFromSheet' -> cria payloads por linha (sem chamar API)
//  - 'getSamplePedido' -> retorna payload estático usado como exemplo

import {
  CodigoDaEmpresa, IdentificadorOperacao, IdentificadorCliente, IdentificadorServico,
  Quantidade, Valor, IdentificadorPrazo, FormaPagamentoEntrada, IdentificadorFormaPagamento,
  NotaCriada, DataEmissao, CFOP, NomeDoServico, Descriminacao1, Descriminacao2, VencimentoFatura
} from './Constantes';

export function main(workbook: any, inputs?: any): any {
  const action = inputs && inputs.action ? inputs.action : 'getSamplePedido';

  if (action === 'buildPedidoVendaFromSheet') {
    const sheet = workbook.getWorksheet('Documento');
    if (!sheet) return { error: 'Worksheet Documento not found' };
    const used = sheet.getUsedRange();
    if (!used) return { payloads: [] };
    const values = used.getValues() as any[][];

    const payloads: any[] = [];
    for (let i = 2; i < values.length; i++) {
      const row = values[i];
      if (!row[NotaCriada] || String(row[NotaCriada]).trim() === '') {
        // Criar observação rica como no .gs
        // Extrair valor numérico da string (ex: "R$ 1.500,00" -> 1500)
        const valorString = String(row[Valor]);
        const valorNumerico = parseFloat(valorString.replace(/[R$\s.,]/g, '').replace(',', '.')) || 0;
        const valorFormatado = new Intl.NumberFormat('pt-BR', {
          style: 'currency',
          currency: 'BRL'
        }).format(valorNumerico);
        
        const observacao = 
          row[NomeDoServico] + '(' + row[Quantidade] + ' X ' + valorFormatado + ') - ' + Valor + '\n\n\n' +
          row[Descriminacao1] + '\n\n' + 
          row[Descriminacao2] + 
          '\n\n Data Vencimento: ' + 
          row[VencimentoFatura];

        const payload: any = {
          CodigoEmpresa: row[CodigoDaEmpresa],
          DataEmissao: row[DataEmissao],
          DataEntrada: row[DataEmissao],
          DataEntrega: row[DataEmissao],
          IdentificadorOperacao: row[IdentificadorOperacao],
          IdentificadorCliente: row[IdentificadorCliente],
          Observacao: observacao,
          Itens: [
            {
              CFOP: row[CFOP],
              COFINS: null,
              PIS: null,
              IdentificadorProduto: row[IdentificadorServico],
              QuantidadePedida: row[Quantidade],
              Valor: valorNumerico,
              ValorUnitario: valorNumerico
            }
          ],
          Status: 'A',
          TipoFrete: 'E'
        };

        if (row[IdentificadorPrazo] && String(row[IdentificadorPrazo]).length > 0) {
          if (String(row[FormaPagamentoEntrada]).toUpperCase() === 'SIM') {
            payload['Prazo'] = {
              Identificador: row[IdentificadorPrazo],
              IdentificadorFormaPagamentoEntrada: row[IdentificadorFormaPagamento]
            };
          } else {
            payload['Prazo'] = {
              Identificador: row[IdentificadorPrazo],
              IdentificadorFormaPagamentoParcelas: row[IdentificadorFormaPagamento]
            };
          }
        }

        payloads.push({ sheetRow: i + 1, payload });
      }
    }

    return { payloads };
  }

  if (action === 'getSamplePedido') {
    // Similar to criarPedidoVenda2 sample payload
    const sample = {
      CodigoEmpresa: 1,
      CodigoEnderecoCobranca: '01',
      CodigoEnderecoEntrega: '01',
      DataEmissao: '2022-08-16 08:00:00',
      DataEntrega: '2022-08-17 00:00:00',
      DataReferenciaPagamento: '2022-08-16 00:00:00',
      IdentificadorCliente: '00A0000063',
      IdentificadorOperacao: '00A000000R',
      IdentificadorPreco: '00A0000005',
      IdentificadorSetor: '00A000000L',
      Itens: [
        {
          DescricaoComplementar: 'Produto Inativo para venda (CdChamada 000001)',
          IdentificadorProduto: '00A00000SQ',
          IdentificadorSetorSaida: '00A000000L',
          PesoBruto: 35,
          PesoLiquido: 10,
          QuantidadePedida: 10,
          Repasses: [],
          Status: 'A',
          Valor: 100,
          ValorUnitario: 10
        },
        {
          CFOP: '',
          DescricaoComplementar: 'Produto ativo para venda (CdChamada 000002)',
          IdentificadorProduto: '00A0000062',
          IdentificadorSetorSaida: '00A000000L',
          OrigemProduto: 0,
          PesoBruto: 35,
          PesoLiquido: 10,
          QuantidadePedida: 10,
          Repasses: [],
          Status: 'A',
          Valor: 150,
          ValorUnitario: 15
        }
      ],
      Observacao: 'Pedido de compra gerado pelo Postman, sem inclusão de mensagem',
      ObservacaoDocumento: 'Pedido de compra gerado pelo BimerWebService',
      Pagamentos: [
        {
          Aliquota: 100,
          AtualizaFinanceiro: true,
          DataReferencia: '2022-08-16 00:00:00',
          Entrada: true,
          IdentificadorFormaPagamento: '00A0000003',
          IdentificadorNaturezaLancamento: '00A0000058',
          Valor: 250
        }
      ],
      Status: 'A',
      TipoFrete: 'E'
    };

    return { sample };
  }

  return { error: 'unknown action' };
}
