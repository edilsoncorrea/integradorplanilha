/**
 * Menu Alternativo para Excel - Substitui menus dinâmicos do Google Sheets
 * Usa botões e Office Scripts para criar interface interativa
 */

export function main(workbook: ExcelScript.Workbook, inputs?: any): any {
  const action = inputs?.action || 'showMenu';

  switch (action) {
    case 'showMenu':
      return createMenuInterface(workbook);
    
    case 'autenticar':
      return executeAuthentication(workbook);
    
    case 'criarPedidos':
      return executeCreatePedidos(workbook);
    
    case 'verificarStatus':
      return executeVerifyStatus(workbook);
    
    case 'limparDados':
      return executeClearData(workbook);
    
    default:
      return { error: 'Ação não reconhecida' };
  }
}

function createMenuInterface(workbook: ExcelScript.Workbook) {
  // Criar aba de controle se não existir
  let controlSheet = workbook.getWorksheet('Controle');
  if (!controlSheet) {
    controlSheet = workbook.addWorksheet('Controle');
  }

  // Limpar conteúdo anterior
  controlSheet.getUsedRange()?.clear();

  // Criar interface de menu
  const menuData = [
    ['🚀 INTEGRADOR DE PLANILHA - MENU PRINCIPAL', '', '', ''],
    ['', '', '', ''],
    ['📋 AÇÕES DISPONÍVEIS:', '', '', ''],
    ['', '', '', ''],
    ['1️⃣ Autenticar na API', 'Clique aqui →', 'autenticar', '✅'],
    ['2️⃣ Criar Pedidos de Venda', 'Clique aqui →', 'criarPedidos', '📦'],
    ['3️⃣ Verificar Status', 'Clique aqui →', 'verificarStatus', '🔍'],
    ['4️⃣ Limpar Dados', 'Clique aqui →', 'limparDados', '🧹'],
    ['', '', '', ''],
    ['📊 STATUS ATUAL:', '', '', ''],
    ['Última Autenticação:', 'Não realizada', '', ''],
    ['Pedidos Processados:', '0', '', ''],
    ['Última Execução:', 'Nunca', '', '']
  ];

  // Inserir dados do menu
  const menuRange = controlSheet.getRangeByIndexes(0, 0, menuData.length, 4);
  menuRange.setValues(menuData);

  // Formatação do menu
  formatMenuInterface(controlSheet);

  return { 
    success: true, 
    message: 'Menu criado na aba "Controle". Use os botões para executar ações.' 
  };
}

function formatMenuInterface(sheet: ExcelScript.Worksheet) {
  // Título principal
  const titleRange = sheet.getRangeByIndexes(0, 0, 1, 4);
  titleRange.getFormat().getFill().setColor('#4472C4');
  titleRange.getFormat().getFont().setColor('white');
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setSize(14);

  // Botões de ação (linhas 4-7)
  for (let i = 4; i <= 7; i++) {
    const buttonRange = sheet.getRangeByIndexes(i, 1, 1, 1);
    buttonRange.getFormat().getFill().setColor('#70AD47');
    buttonRange.getFormat().getFont().setColor('white');
    buttonRange.getFormat().getFont().setBold(true);
  }

  // Seção de status
  const statusRange = sheet.getRangeByIndexes(9, 0, 1, 4);
  statusRange.getFormat().getFill().setColor('#FFC000');
  statusRange.getFormat().getFont().setBold(true);

  // Ajustar largura das colunas
  sheet.getColumn(0).setWidth(200);
  sheet.getColumn(1).setWidth(150);
  sheet.getColumn(2).setWidth(100);
  sheet.getColumn(3).setWidth(50);
}

function executeAuthentication(workbook: ExcelScript.Workbook) {
  // Importar e executar autenticação
  const { main: authMain } = require('./Autenticacao');
  
  const result = authMain(workbook, {
    action: 'buildAuthPayload',
    host: 'https://087344bimerapi.alterdata.cloud',
    username: 'bimerapi',
    senha: '123456',
    nonce: '123456789'
  });

  // Atualizar status na aba Controle
  updateStatus(workbook, 'auth', new Date().toLocaleString());
  
  return {
    success: true,
    message: '✅ Autenticação configurada! Verifique a aba Controle.',
    data: result
  };
}

function executeCreatePedidos(workbook: ExcelScript.Workbook) {
  // Importar e executar criação de pedidos
  const { main: pedidoMain } = require('./PedidoDeVenda');
  
  const result = pedidoMain(workbook, { 
    action: 'buildPedidoVendaFromSheet' 
  });

  // Atualizar status
  updateStatus(workbook, 'pedidos', result.payloads?.length || 0);
  updateStatus(workbook, 'execucao', new Date().toLocaleString());
  
  return {
    success: true,
    message: `✅ ${result.payloads?.length || 0} pedidos processados! Verifique a aba Documento.`,
    data: result
  };
}

function executeVerifyStatus(workbook: ExcelScript.Workbook) {
  const documentoSheet = workbook.getWorksheet('Documento');
  if (!documentoSheet) {
    return { success: false, message: '❌ Aba Documento não encontrada' };
  }

  const usedRange = documentoSheet.getUsedRange();
  if (!usedRange) {
    return { success: false, message: '❌ Nenhum dado encontrado' };
  }

  const values = usedRange.getValues();
  let totalLinhas = values.length - 2; // Excluir headers
  let pedidosCriados = 0;
  let pedidosPendentes = 0;

  for (let i = 2; i < values.length; i++) {
    const notaCriada = values[i][21]; // Coluna NotaCriada
    if (notaCriada && String(notaCriada).trim() !== '') {
      pedidosCriados++;
    } else {
      pedidosPendentes++;
    }
  }

  const statusMessage = `📊 STATUS:\n` +
    `Total de linhas: ${totalLinhas}\n` +
    `Pedidos criados: ${pedidosCriados}\n` +
    `Pedidos pendentes: ${pedidosPendentes}`;

  return {
    success: true,
    message: statusMessage,
    data: { totalLinhas, pedidosCriados, pedidosPendentes }
  };
}

function executeClearData(workbook: ExcelScript.Workbook) {
  const documentoSheet = workbook.getWorksheet('Documento');
  if (!documentoSheet) {
    return { success: false, message: '❌ Aba Documento não encontrada' };
  }

  const usedRange = documentoSheet.getUsedRange();
  if (!usedRange) {
    return { success: false, message: '❌ Nenhum dado para limpar' };
  }

  const values = usedRange.getValues();
  let limpezas = 0;

  // Limpar colunas NotaCriada e RetornoAPI
  for (let i = 2; i < values.length; i++) {
    documentoSheet.getRangeByIndexes(i, 21, 1, 1).setValue(''); // NotaCriada
    documentoSheet.getRangeByIndexes(i, 22, 1, 1).setValue(''); // RetornoAPI
    limpezas++;
  }

  // Atualizar status
  updateStatus(workbook, 'pedidos', 0);
  updateStatus(workbook, 'execucao', 'Dados limpos em ' + new Date().toLocaleString());

  return {
    success: true,
    message: `🧹 ${limpezas} linhas limpas com sucesso!`
  };
}

function updateStatus(workbook: ExcelScript.Workbook, type: string, value: any) {
  const controlSheet = workbook.getWorksheet('Controle');
  if (!controlSheet) return;

  switch (type) {
    case 'auth':
      controlSheet.getRangeByIndexes(10, 1, 1, 1).setValue(value);
      break;
    case 'pedidos':
      controlSheet.getRangeByIndexes(11, 1, 1, 1).setValue(value);
      break;
    case 'execucao':
      controlSheet.getRangeByIndexes(12, 1, 1, 1).setValue(value);
      break;
  }
}