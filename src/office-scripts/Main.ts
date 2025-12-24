// Main.ts
// Conversão parcial de Main.gs. Office Scripts não possui onOpen ou menus da mesma forma que GAS.
// Este arquivo fornece ações que podem ser chamadas por Power Automate ou manualmente:
//  - 'authPayload' -> direciona para uso do script Autenticacao (buildAuthPayload)
//  - 'buildPedido' -> direciona para uso do PedidoDeVenda script
//  - 'showSpinnerNote' -> informação sobre o Spinner.html (não aplicável no Office Scripts)

export function main(workbook: any, inputs?: any): any {
  const action = inputs && inputs.action ? inputs.action : 'help';

  if (action === 'help') {
    return {
      note: 'Office Scripts não suporta onOpen menus. Use botões no Excel Online ou Power Automate to run scripts.'
    };
  }

  if (action === 'authPayload') {
    return { note: 'Execute o script Autenticacao (action buildAuthPayload) para gerar URL/payload do token.' };
  }

  if (action === 'buildPedido') {
    return { note: 'Execute o script PedidoDeVenda (action buildPedidoVendaFromSheet) para gerar payloads.' };
  }

  if (action === 'showSpinnerNote') {
    return { note: 'Spinner.html não é aplicável. Use ação do Power Automate ou mensagem de progresso no Flow.' };
  }

  return { error: 'unknown action' };
}
