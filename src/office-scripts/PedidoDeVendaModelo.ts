// PedidoDeVendaModelo.ts
// Conversão de PedidoDeVendaModelo.gs — exporta payload de exemplo usado para testes

export function getPedidoDeVendaModelo() {
  return {
    CodigoEmpresa: 1,
    CodigoEnderecoCobranca: "01",
    CodigoEnderecoEntrega: "01",
    DataEmissao: "2022-18-08 00:00:00",
    DataEntrega: "2022-18-08 00:00:00",
    DataReferenciaPagamento: "2022-18-08 00:00:00",
    IdentificadorCliente: "00A0000063",
    IdentificadorOperacao: "00A000000R",
    IdentificadorPreco: "00A0000005",
    IdentificadorSetor: "00A000000L",
    Itens: [
      {
        DescricaoComplementar: "Produto Inativo para venda (CdChamada 000001)",
        IdentificadorProduto: "00A00000SQ",
        IdentificadorSetorSaida: "00A000000L",
        PesoBruto: 35,
        PesoLiquido: 10,
        QuantidadePedida: 10,
        Repasses: [],
        Status: "A",
        Valor: 100.00,
        ValorUnitario: 10.00
      },
      {
        CFOP: "",
        DescricaoComplementar: "Produto ativo para venda (CdChamada 000002)",
        IdentificadorProduto: "00A0000062",
        IdentificadorSetorSaida: "00A000000L",
        OrigemProduto: 0,
        PesoBruto: 35,
        PesoLiquido: 10,
        QuantidadePedida: 10,
        Repasses: [],
        Status: "A",
        Valor: 150.00,
        ValorUnitario: 15.00
      }
    ],
    Observacao: "Pedido de compra gerado pelo Postman, sem inclusão de mensagem",
    ObservacaoDocumento: "Pedido de compra gerado pelo BimerWebService",
    Pagamentos: [
      {
        Aliquota: 100,
        AtualizaFinanceiro: true,
        DataReferencia: "2022-18-08 00:00:00",
        Entrada: true,
        IdentificadorFormaPagamento: "00A0000003",
        IdentificadorNaturezaLancamento: "00A0000058",
        Valor: 250.00
      }
    ],
    Status: "A",
    TipoFrete: "E"
  };
}
