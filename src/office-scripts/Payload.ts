// Payload.ts
// Conversão de Payload.gs para Office Scripts (TypeScript)

export function getPayloadSample() {
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
        Valor: 100.0,
        ValorUnitario: 10.0
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
        Valor: 150.0,
        ValorUnitario: 15.0
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
        Valor: 250.0
      }
    ],
    Status: "A",
    TipoFrete: "E"
  };
}
