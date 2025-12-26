/**
 * ═══════════════════════════════════════════════════════════════════════════
 * INTEGRADOR COMPLETO - OFFICE SCRIPTS PARA EXCEL ONLINE
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Script consolidado que integra todas as funcionalidades para criação de
 * Pedidos de Venda e Documentos a partir de planilhas Excel, com integração
 * à API Bimer.
 * 
 * FUNCIONALIDADES PRINCIPAIS:
 * ────────────────────────────────────────────────────────────────────────────
 * 1. AUTENTICAÇÃO: Gera payload MD5 para autenticação na API Bimer
 * 2. VALIDAÇÃO: Valida dados da planilha e busca identificadores na API
 * 3. PEDIDOS: Cria payloads de pedidos de venda a partir dos dados da planilha
 * 4. DOCUMENTOS: Cria payloads de documentos fiscais a partir dos dados
 * 5. RESULTADOS: Aplica resultados das APIs de volta na planilha
 * 6. ⭐ EXECUÇÃO COMPLETA: Executa todo o fluxo com chamadas HTTP diretas
 * 
 * NOVO - EXECUÇÃO DIRETA:
 * ────────────────────────────────────────────────────────────────────────────
 * Este script AGORA FAZ chamadas HTTP diretamente usando fetch()!
 * Use action='executarCompleto' para processar toda a planilha automaticamente:
 *   1. Autentica na API Bimer
 *   2. Valida e busca identificadores faltantes
 *   3. Cria documentos via API
 *   4. Atualiza resultados na planilha
 * 
 * USO SIMPLIFICADO (RECOMENDADO):
 * ────────────────────────────────────────────────────────────────────────────
 * Basta executar o script sem parâmetros ou com:
 *   { action: 'executarCompleto' }
 * 
 * O script processará automaticamente todas as linhas da planilha que ainda
 * não têm "Nota Criada" preenchida.
 * 
 * USO AVANÇADO (Para Power Automate):
 * ────────────────────────────────────────────────────────────────────────────
 * Use o parâmetro 'action' em inputs para escolher operações específicas:
 * 
 * AUTENTICAÇÃO:
 *   - action: 'buildAuthPayload' -> Gera URL e payload para obter token
 *   - action: 'hash' -> Calcula MD5 de um valor
 * 
 * VALIDAÇÃO:
 *   - action: 'buildValidationQueries' -> Lista consultas GET necessárias
 *   - action: 'applyValidationResults' -> Aplica IDs na planilha
 * 
 * PEDIDOS:
 *   - action: 'buildPedidos' -> Gera payloads de pedidos de venda
 * 
 * DOCUMENTOS:
 *   - action: 'buildDocumentos' -> Gera payloads de documentos fiscais
 * 
 * RESULTADOS:
 *   - action: 'applyResults' -> Escreve resultados na planilha
 * 
 * CONFIGURAÇÃO:
 * ────────────────────────────────────────────────────────────────────────────
 * Credenciais padrão (altere se necessário):
 *   - Host: https://homologacaowisepcp.alterdata.com.br/BimerApi
 *   - Username: supervisor
 *   - Senha: Senhas123
 * 
 * Para usar credenciais diferentes, passe nos inputs:
 *   { 
 *     action: 'executarCompleto',
 *     host: 'sua-url',
 *     username: 'seu-usuario',
 *     senha: 'sua-senha'
 *   }
 * 
 * ESTRUTURA DA PLANILHA:
 * ────────────────────────────────────────────────────────────────────────────
 * Nome da aba: "Documento"
 * Colunas esperadas (índice 0-based):
 *   0:  Código da Empresa
 *   1:  Código Cliente
 *   2:  Nome do Cliente
 *   3:  Identificador Cliente (preenchido automaticamente se vazio)
 *   4:  Código da Operação
 *   5:  Identificador Operação (preenchido automaticamente se vazio)
 *   6:  CFOP
 *   7:  Código do Serviço
 *   8:  Identificador Serviço (preenchido automaticamente se vazio)
 *   9:  Nome do Serviço
 *   10: Quantidade
 *   11: Valor
 *   12: Discriminação 1
 *   13: Discriminação 2
 *   14: Código Prazo
 *   15: Identificador Prazo
 *   16: Forma Pagamento Entrada
 *   17: Código da Forma de Pagamento
 *   18: Identificador Forma Pagamento (preenchido automaticamente se vazio)
 *   19: Data Emissão
 *   20: Vencimento Fatura
 *   21: Nota Criada (✅ preenchido pelo script: "Sim" ou "Não")
 *   22: Retorno API (✅ preenchido pelo script: identificador ou mensagem erro)
 *   23: Log de Execução (✅ preenchido pelo script: detalhes do processamento)
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════
// CONSTANTES - Índices das colunas (0-based)
// ═══════════════════════════════════════════════════════════════════════════

const HOST = "https://homologacaowisepcp.alterdata.com.br/BimerApi";

const CodigoDaEmpresa = 0;
const CodigoCliente = 1;
const NomeDoCliente = 2;
const IdentificadorCliente = 3; 
const CodigoDaOperacao = 4; 
const IdentificadorOperacao = 5; 
const CFOP = 6; 
const CodigoDoServico = 7; 
const IdentificadorServico = 8; 
const NomeDoServico = 9;
const Quantidade = 10;
const Valor = 11; 
const Descriminacao1 = 12; 
const Descriminacao2 = 13; 
const Codigoprazo = 14;
const IdentificadorPrazo = 15;
const FormaPagamentoEntrada = 16;
const CodigoDaFormaDePagamento = 17;
const IdentificadorFormaPagamento = 18;
const DataEmissao = 19;
const VencimentoFatura = 20; 
const NotaCriada = 21;
const RetornoAPI = 22;
const LogExecucao = 23;

// ═══════════════════════════════════════════════════════════════════════════
// FUNÇÕES DE LOGGING E DEBUG
// ═══════════════════════════════════════════════════════════════════════════

interface LogEntry {
  timestamp: string;
  nivel: 'INFO' | 'WARNING' | 'ERROR' | 'DEBUG';
  mensagem: string;
  contexto?: string;
}

class Logger {
  private logs: LogEntry[] = [];
  private habilitarConsole: boolean = true;

  constructor(habilitarConsole: boolean = true) {
    this.habilitarConsole = habilitarConsole;
  }

  private formatTimestamp(): string {
    const now = new Date();
    return `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}:${now.getSeconds().toString().padStart(2, '0')}.${now.getMilliseconds().toString().padStart(3, '0')}`;
  }

  log(nivel: 'INFO' | 'WARNING' | 'ERROR' | 'DEBUG', mensagem: string, contexto?: string) {
    const entry: LogEntry = {
      timestamp: this.formatTimestamp(),
      nivel,
      mensagem,
      contexto
    };

    this.logs.push(entry);

    if (this.habilitarConsole) {
      const emoji = nivel === 'INFO' ? 'ℹ️' : nivel === 'WARNING' ? '⚠️' : nivel === 'ERROR' ? '❌' : '🔍';
      const msg = contexto ? `[${entry.timestamp}] ${emoji} ${nivel}: ${mensagem} (${contexto})` : `[${entry.timestamp}] ${emoji} ${nivel}: ${mensagem}`;
      console.log(msg);
    }
  }

  info(mensagem: string, contexto?: string) {
    this.log('INFO', mensagem, contexto);
  }

  warning(mensagem: string, contexto?: string) {
    this.log('WARNING', mensagem, contexto);
  }

  error(mensagem: string, contexto?: string) {
    this.log('ERROR', mensagem, contexto);
  }

  debug(mensagem: string, contexto?: string) {
    this.log('DEBUG', mensagem, contexto);
  }

  getLogs(): LogEntry[] {
    return this.logs;
  }

  getLogsFormatted(): string {
    return this.logs.map(l => `[${l.timestamp}] ${l.nivel}: ${l.mensagem}${l.contexto ? ' (' + l.contexto + ')' : ''}`).join('\n');
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// PONTO DE ENTRADA PRINCIPAL
// ═══════════════════════════════════════════════════════════════════════════

async function main(workbook: ExcelScript.Workbook, inputs?: {
  action?: string;
  host?: string;
  username?: string;
  senha?: string;
  nonce?: string;
  value?: string;
  results?: object[];
  executeAPI?: boolean;
}) {
  const action = inputs?.action || 'executarCompleto';

  // EXECUÇÃO COMPLETA COM API
  if (action === 'executarCompleto') return await executarFluxoCompleto(workbook, inputs);

  // AUTENTICAÇÃO
  if (action === 'buildAuthPayload') return buildAuthPayload(inputs);
  if (action === 'hash') return hashValue(inputs);

  // VALIDAÇÃO
  if (action === 'buildValidationQueries') return buildValidationQueries(workbook);
  if (action === 'applyValidationResults') return applyValidationResults(workbook, inputs);

  // PEDIDOS DE VENDA
  if (action === 'buildPedidos') return buildPedidos(workbook);

  // DOCUMENTOS FISCAIS
  if (action === 'buildDocumentos') return buildDocumentos(workbook);

  // APLICAR RESULTADOS
  if (action === 'applyResults') return applyResults(workbook, inputs);

  // AJUDA
  if (action === 'help') {
    const help = {
      message: 'INTEGRADOR COMPLETO - Office Scripts para Excel Online',
      actions: {
        execucao: ['executarCompleto - Executa fluxo completo com chamadas HTTP'],
        autenticacao: ['buildAuthPayload', 'hash'],
        validacao: ['buildValidationQueries', 'applyValidationResults'],
        pedidos: ['buildPedidos'],
        documentos: ['buildDocumentos'],
        resultados: ['applyResults']
      },
      usage: 'Passe inputs.action com uma das ações listadas acima',
      exemplo: '{ action: "executarCompleto" }',
      recomendado: 'Use action="executarCompleto" para processar tudo automaticamente'
    };
    return help;
  }

  return { error: `Ação desconhecida: ${action}` };
}

// ═══════════════════════════════════════════════════════════════════════════
// SEÇÃO 1: AUTENTICAÇÃO
// ═══════════════════════════════════════════════════════════════════════════

function buildAuthPayload(inputs?: { host?: string; username?: string; senha?: string; nonce?: string }) {
  const host = (inputs?.host as string) || HOST;
  const username = (inputs?.username as string) || 'supervisor';
  const senha = (inputs?.senha as string) || 'Senhas123';
  const nonce = (inputs?.nonce as string) || '123456789';

  const password = md5(username + nonce + senha);

  const payload = {
    client_id: 'IntegracaoBimer.js',
    username: username,
    password: password,
    grant_type: 'password',
    nonce: nonce
  };

  const url = host.endsWith('/') ? host + 'oauth/token' : host + '/oauth/token';

  return {
    url: url,
    method: 'POST',
    payload: payload,
    note: 'Use este payload no Power Automate para fazer POST e obter access_token'
  };
}

function hashValue(inputs?: { value?: string }) {
  const value = (inputs?.value as string) || '';
  return { md5: md5(value) };
}

// ═══════════════════════════════════════════════════════════════════════════
// SEÇÃO 1B: EXECUÇÃO COMPLETA COM CHAMADAS HTTP
// ═══════════════════════════════════════════════════════════════════════════

async function executarFluxoCompleto(workbook: ExcelScript.Workbook, inputs?: {
  host?: string;
  username?: string;
  senha?: string;
  nonce?: string;
}) {
  const logger = new Logger(true);
  const host = (inputs?.host as string) || HOST;
  const username = (inputs?.username as string) || 'supervisor';
  const senha = (inputs?.senha as string) || 'Senhas123';

  let totalProcessados = 0;
  let totalSucesso = 0;
  let totalErro = 0;

  try {
    logger.info('═══════════════════════════════════════════════════════');
    logger.info('INICIANDO EXECUÇÃO COMPLETA DO INTEGRADOR BIMER');
    logger.info('═══════════════════════════════════════════════════════');
    logger.debug(`Host: ${host}`);
    logger.debug(`Usuário: ${username}`);

    // 1. AUTENTICAR
    logger.info('🔐 Iniciando autenticação na API Bimer...');
    const inicioAuth = Date.now();
    const token = await autenticarAPI(host, username, senha, logger);
    const tempoAuth = Date.now() - inicioAuth;
    logger.info(`✅ Token obtido com sucesso em ${tempoAuth}ms`);
    logger.debug(`Token (primeiros 20 chars): ${token.substring(0, 20)}...`);

    // 2. VALIDAR PLANILHA (buscar identificadores faltantes)
    logger.info('🔍 Iniciando validação de identificadores na planilha...');
    const inicioValidacao = Date.now();
    const validados = await validarPlanilha(workbook, token, host, logger);
    const tempoValidacao = Date.now() - inicioValidacao;
    logger.info(`✅ ${validados} identificador(es) validado(s) em ${tempoValidacao}ms`);

    // 3. CRIAR DOCUMENTOS
    logger.info('📄 Iniciando criação de documentos na API...');
    const inicioDocumentos = Date.now();
    const resultados = await criarDocumentosAPI(workbook, token, host, logger);
    const tempoDocumentos = Date.now() - inicioDocumentos;
    
    totalProcessados = resultados.length;
    totalSucesso = resultados.filter(r => r.sucesso).length;
    totalErro = resultados.filter(r => !r.sucesso).length;

    logger.info(`✅ ${totalSucesso} documento(s) criado(s) em ${tempoDocumentos}ms`);
    if (totalErro > 0) logger.warning(`${totalErro} erro(s) encontrado(s)`);

    const tempoTotal = tempoAuth + tempoValidacao + tempoDocumentos;
    logger.info('═══════════════════════════════════════════════════════');
    logger.info(`PROCESSO CONCLUÍDO: ${totalSucesso}/${totalProcessados} documentos criados`);
    logger.info(`Tempo total de execução: ${tempoTotal}ms`);
    logger.info('═══════════════════════════════════════════════════════');

    return {
      sucesso: true,
      totalProcessados,
      totalSucesso,
      totalErro,
      tempoExecucao: tempoTotal,
      logs: logger.getLogs(),
      logFormatado: logger.getLogsFormatted(),
      mensagem: `Processo concluído: ${totalSucesso}/${totalProcessados} documentos criados em ${tempoTotal}ms`
    };

  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logger.error(`ERRO FATAL: ${errorMessage}`);
    
    // Diagnóstico adicional para erros de fetch
    if (errorMessage.includes('Failed to fetch') || errorMessage.includes('Falha na conexão')) {
      logger.error('POSSÍVEIS CAUSAS:');
      logger.error('1. API Bimer está fora do ar ou inacessível');
      logger.error('2. URL da API está incorreta');
      logger.error('3. Problema de firewall ou rede');
      logger.error('4. Política CORS bloqueando requisição do Office Scripts');
      logger.error('SOLUÇÃO: Teste a URL no navegador: https://homologacaowisepcp.alterdata.com.br/BimerApi/oauth/token');
    }
    
    return {
      sucesso: false,
      erro: errorMessage,
      logs: logger.getLogs(),
      logFormatado: logger.getLogsFormatted()
    };
  }
}

async function autenticarAPI(host: string, username: string, senha: string, logger?: Logger): Promise<string> {
  const nonce = Date.now().toString();
  const password = md5(username + nonce + senha);

  if (logger) {
    logger.debug(`Nonce gerado: ${nonce}`);
    logger.debug(`Hash MD5 calculado`);
  }

  const params = new URLSearchParams({
    client_id: 'IntegracaoBimer.js',
    username: username,
    password: password,
    grant_type: 'password',
    nonce: nonce
  });

  const url = host.endsWith('/') ? `${host}oauth/token` : `${host}/oauth/token`;
  
  if (logger) logger.debug(`URL de autenticação: ${url}`);

  let response: Response;
  try {
    response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: params.toString()
    });
  } catch (fetchError) {
    const errorMsg = `Falha na conexão com a API. Verifique: 1) URL está correta? 2) API está acessível? 3) Sem bloqueio de CORS? Erro: ${fetchError}`;
    if (logger) logger.error(errorMsg);
    throw new Error(errorMsg);
  }

  if (logger) logger.debug(`Status da resposta: ${response.status}`);

  if (!response.ok) {
    const errorMsg = `Erro na autenticação: ${response.status} ${response.statusText}`;
    if (logger) logger.error(errorMsg);
    throw new Error(errorMsg);
  }

  const data: { access_token: string } = await response.json() as { access_token: string };
  if (logger) logger.debug('Token de acesso recebido com sucesso');
  return data.access_token;
}

async function validarPlanilha(workbook: ExcelScript.Workbook, token: string, host: string, logger?: Logger): Promise<number> {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) {
    const erro = 'Planilha "Documento" não encontrada';
    if (logger) logger.error(erro);
    throw new Error(erro);
  }

  const used = sheet.getUsedRange();
  if (!used) {
    if (logger) logger.warning('Nenhuma área usada na planilha');
    return 0;
  }

  const values = used.getValues();
  if (logger) logger.debug(`Total de linhas na planilha: ${values.length}`);
  let validados = 0;

  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const notaCriada = row[NotaCriada];

    // Pula linhas já processadas
    if (notaCriada && String(notaCriada).trim() !== '') continue;

    // Validar Cliente
    if (!row[IdentificadorCliente] || String(row[IdentificadorCliente]).trim() === '') {
      const codigoCliente = row[CodigoCliente];
      if (codigoCliente && typeof codigoCliente !== 'boolean') {
        // Try-catch necessário: continua processamento mesmo com erro em linha específica
        try {
          if (logger) logger.debug(`Buscando cliente código ${codigoCliente}`, `Linha ${i + 1}`);
          const cliente: { Identificador: string; Nome: string } = await buscarPessoaPorCodigo(host, token, String(codigoCliente)) as { Identificador: string; Nome: string };
          sheet.getRangeByIndexes(i, IdentificadorCliente, 1, 1).setValue(cliente.Identificador);
          sheet.getRangeByIndexes(i, NomeDoCliente, 1, 1).setValue(cliente.Nome || '');
          validados++;
          if (logger) logger.info(`Cliente ${codigoCliente} encontrado: ${cliente.Nome}`, `Linha ${i + 1}`);
        } catch (error) {
          if (logger) logger.error(`Erro ao buscar cliente ${codigoCliente}: ${error}`, `Linha ${i + 1}`);
        }
      }
    }

    // Validar Forma de Pagamento
    if (!row[IdentificadorFormaPagamento] || String(row[IdentificadorFormaPagamento]).trim() === '') {
      const codigoFormaPagamento = row[CodigoDaFormaDePagamento];
      if (codigoFormaPagamento && typeof codigoFormaPagamento !== 'boolean') {
        // Try-catch necessário: continua processamento mesmo com erro em linha específica
        try {
          if (logger) logger.debug(`Buscando forma de pagamento código ${codigoFormaPagamento}`, `Linha ${i + 1}`);
          const formas: { Codigo: string | number; Identificador: string }[] = await buscarFormasPagamento(host, token) as { Codigo: string | number; Identificador: string }[];
          const forma: { Codigo: string | number; Identificador: string } | undefined = formas.find((f: { Codigo: string | number | boolean }) => 
            String(f.Codigo) === String(codigoFormaPagamento)
          );
          if (forma) {
            sheet.getRangeByIndexes(i, IdentificadorFormaPagamento, 1, 1).setValue(forma.Identificador);
            validados++;
            if (logger) logger.info(`Forma de pagamento ${codigoFormaPagamento} encontrada`, `Linha ${i + 1}`);
          } else {
            if (logger) logger.warning(`Forma de pagamento ${codigoFormaPagamento} não encontrada`, `Linha ${i + 1}`);
          }
        } catch (error) {
          if (logger) logger.error(`Erro ao buscar forma de pagamento: ${error}`, `Linha ${i + 1}`);
        }
      }
    }

    // Validar Operação
    if (!row[IdentificadorOperacao] || String(row[IdentificadorOperacao]).trim() === '') {
      const codigoOperacao = row[CodigoDaOperacao];
      if (codigoOperacao && typeof codigoOperacao !== 'boolean') {
        // Try-catch necessário: continua processamento mesmo com erro em linha específica
        try {
          if (logger) logger.debug(`Buscando operação código ${codigoOperacao}`, `Linha ${i + 1}`);
          const operacoes: { Codigo: string | number; Identificador: string }[] = await buscarOperacoes(host, token) as { Codigo: string | number; Identificador: string }[];
          const operacao: { Codigo: string | number; Identificador: string } | undefined = operacoes.find((o: { Codigo: string | number | boolean }) => 
            String(o.Codigo) === String(codigoOperacao)
          );
          if (operacao) {
            sheet.getRangeByIndexes(i, IdentificadorOperacao, 1, 1).setValue(operacao.Identificador);
            validados++;
            if (logger) logger.info(`Operação ${codigoOperacao} encontrada`, `Linha ${i + 1}`);
          } else {
            if (logger) logger.warning(`Operação ${codigoOperacao} não encontrada`, `Linha ${i + 1}`);
          }
        } catch (error) {
          if (logger) logger.error(`Erro ao buscar operação: ${error}`, `Linha ${i + 1}`);
        }
      }
    }

    // Validar Produto/Serviço
    if (!row[IdentificadorServico] || String(row[IdentificadorServico]).trim() === '') {
      const codigoServico = row[CodigoDoServico];
      if (codigoServico && typeof codigoServico !== 'boolean') {
        // Try-catch necessário: continua processamento mesmo com erro em linha específica
        try {
          if (logger) logger.debug(`Buscando produto código ${codigoServico}`, `Linha ${i + 1}`);
          const produtos: { Identificador: string }[] = await buscarProdutos(host, token, String(codigoServico)) as { Identificador: string }[];
          if (produtos.length > 0) {
            sheet.getRangeByIndexes(i, IdentificadorServico, 1, 1).setValue(produtos[0].Identificador);
            validados++;
            if (logger) logger.info(`Produto ${codigoServico} encontrado`, `Linha ${i + 1}`);
          } else {
            if (logger) logger.warning(`Produto ${codigoServico} não encontrado`, `Linha ${i + 1}`);
          }
        } catch (error) {
          if (logger) logger.error(`Erro ao buscar produto: ${error}`, `Linha ${i + 1}`);
        }
      }
    }
  }

  return validados;
}

async function criarDocumentosAPI(workbook: ExcelScript.Workbook, token: string, host: string, logger?: Logger) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) {
    const erro = 'Planilha "Documento" não encontrada';
    if (logger) logger.error(erro);
    throw new Error(erro);
  }

  const used = sheet.getUsedRange();
  if (!used) {
    if (logger) logger.warning('Nenhuma área usada na planilha');
    return [];
  }

  const values = used.getValues();
  const resultados: { linha: number; sucesso: boolean; mensagem: string }[] = [];
  let processadas = 0;

  if (logger) logger.debug(`Processando ${values.length - 2} linha(s) de dados`);

  for (let i = 2; i < values.length; i++) {
    processadas++;
    const row = values[i];
    const notaCriada = row[NotaCriada];

    // Pula linhas já processadas
    if (notaCriada && String(notaCriada).trim() !== '') {
      if (logger) logger.debug(`Linha ${i + 1} já processada, pulando`);
      continue;
    }

    if (logger) logger.info(`Processando linha ${i + 1}...`);

    // Validar campos obrigatórios
    if (!row[IdentificadorCliente] || !row[IdentificadorOperacao] || !row[IdentificadorServico]) {
      const erro = 'Campos obrigatórios faltando (Cliente, Operação ou Serviço)';
      if (logger) logger.error(erro, `Linha ${i + 1}`);
      sheet.getRangeByIndexes(i, NotaCriada, 1, 1).setValue('Não');
      sheet.getRangeByIndexes(i, RetornoAPI, 1, 1).setValue(erro);
      sheet.getRangeByIndexes(i, LogExecucao, 1, 1).setValue(`[ERRO] ${erro}`);
      resultados.push({ linha: i + 1, sucesso: false, mensagem: erro });
      continue;
    }

    // Construir payload
    if (logger) logger.debug(`Construindo payload do documento`, `Linha ${i + 1}`);
    const valorString = String(row[Valor]);
    const valorNumerico = parseValor(valorString);
    const valorFormatado = formatCurrency(valorNumerico);

    if (logger) logger.debug(`Valor: ${valorFormatado} (${valorNumerico})`, `Linha ${i + 1}`);

    const observacao = 
      `${row[NomeDoServico]} (${row[Quantidade]} X ${valorFormatado}) - ${valorNumerico}\n\n\n` +
      `${row[Descriminacao1]}\n\n${row[Descriminacao2]}` +
      `\n\n Data Vencimento: ${row[VencimentoFatura]}`;

    const payload = {
      StatusNotaFiscalEletronica: 'A',
      TipoDocumento: 'S',
      TipoPagamento: '0',
      CodigoEmpresa: row[CodigoDaEmpresa],
      DataEmissao: row[DataEmissao],
      DataReferencia: row[DataEmissao],
      DataReferenciaPagamento: row[DataEmissao],
      IdentificadorOperacao: row[IdentificadorOperacao],
      IdentificadorPessoa: row[IdentificadorCliente],
      Observacao: observacao,
      Itens: [
        {
          CFOP: row[CFOP],
          IdentificadorProduto: row[IdentificadorServico],
          Quantidade: row[Quantidade],
          ValorUnitario: valorNumerico
        }
      ],
      Pagamentos: [
        {
          Aliquota: 100,
          AliquotaConvenio: 10,
          DataVencimento: row[DataEmissao],
          IdentificadorFormaPagamento: row[IdentificadorFormaPagamento],
          Valor: valorNumerico
        }
      ]
    };

    // Criar documento via API
    // Try-catch necessário: captura erros de API e continua processando outras linhas
    try {
      const url = host.endsWith('/') ? `${host}api/documentos` : `${host}/api/documentos`;
      if (logger) logger.debug(`POST para ${url}`, `Linha ${i + 1}`);
      
      const inicioReq = Date.now();
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      });
      const tempoReq = Date.now() - inicioReq;

      if (logger) logger.debug(`Resposta recebida em ${tempoReq}ms - Status: ${response.status}`, `Linha ${i + 1}`);

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`${response.status}: ${errorText}`);
      }

      const data: { ListaObjetos?: { Identificador: string }[] } = await response.json() as { ListaObjetos?: { Identificador: string }[] };
      const identificador: string = data.ListaObjetos?.[0]?.Identificador || 'Criado';

      if (logger) logger.info(`✅ Documento criado: ${identificador}`, `Linha ${i + 1}`);

      sheet.getRangeByIndexes(i, NotaCriada, 1, 1).setValue('Sim');
      sheet.getRangeByIndexes(i, RetornoAPI, 1, 1).setValue(identificador);
      sheet.getRangeByIndexes(i, LogExecucao, 1, 1).setValue(`[OK] Criado em ${tempoReq}ms - ID: ${identificador}`);
      
      resultados.push({ linha: i + 1, sucesso: true, mensagem: identificador });

    } catch (error) {
      const mensagemErro = String(error);
      if (logger) logger.error(`❌ Falha ao criar documento: ${mensagemErro}`, `Linha ${i + 1}`);
      
      sheet.getRangeByIndexes(i, NotaCriada, 1, 1).setValue('Não');
      sheet.getRangeByIndexes(i, RetornoAPI, 1, 1).setValue(mensagemErro);
      sheet.getRangeByIndexes(i, LogExecucao, 1, 1).setValue(`[ERRO] ${mensagemErro}`);
      
      resultados.push({ linha: i + 1, sucesso: false, mensagem: mensagemErro });
    }
  }

  if (logger) logger.info(`Total processado: ${processadas} linha(s)`);

  return resultados;
}

// Funções auxiliares de API

async function buscarPessoaPorCodigo(host: string, token: string, codigo: string): Promise<{ Identificador: string; Nome: string }> {
  const url = host.endsWith('/') 
    ? `${host}api/pessoas/codigo/${codigo}` 
    : `${host}/api/pessoas/codigo/${codigo}`;

  let response: Response;
  try {
    response = await fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });
  } catch (error) {
    throw new Error(`Falha na conexão ao buscar pessoa: ${error}`);
  }

  if (!response.ok) {
    throw new Error(`Erro ao buscar pessoa: ${response.status}`);
  }

  return await response.json() as { Identificador: string; Nome: string };
}

async function buscarFormasPagamento(host: string, token: string): Promise<{ Codigo: string | number; Identificador: string }[]> {
  const url = host.endsWith('/') 
    ? `${host}api/formasPagamento` 
    : `${host}/api/formasPagamento`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    throw new Error(`Erro ao buscar formas de pagamento: ${response.status}`);
  }

  return await response.json() as { Codigo: string | number; Identificador: string }[];
}

async function buscarOperacoes(host: string, token: string): Promise<{ Codigo: string | number; Identificador: string }[]> {
  const url = host.endsWith('/') 
    ? `${host}api/operacoes` 
    : `${host}/api/operacoes`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    throw new Error(`Erro ao buscar operações: ${response.status}`);
  }

  return await response.json() as { Codigo: string | number; Identificador: string }[];
}

async function buscarProdutos(host: string, token: string, codigo: string): Promise<{ Identificador: string }[]> {
  const url = host.endsWith('/') 
    ? `${host}api/produtos?codigo=${codigo}` 
    : `${host}/api/produtos?codigo=${codigo}`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    throw new Error(`Erro ao buscar produtos: ${response.status}`);
  }

  return await response.json() as { Identificador: string }[];
}

// ═══════════════════════════════════════════════════════════════════════════
// SEÇÃO 2: VALIDAÇÃO DA PLANILHA
// ═══════════════════════════════════════════════════════════════════════════

function buildValidationQueries(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const used = sheet.getUsedRange();
  if (!used) return { queries: [], total: 0, note: 'Nenhuma área usada na planilha. Sem consultas a executar.' };
  const values: (string | number | boolean)[][] = used.getValues();

  const queries: { sheetRow: number; method: string; endpoint: string; field: string; codigo: string | number }[] = [];

  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const notaCriada = row[NotaCriada];
    
    // Pula linhas que já têm nota criada
    if (notaCriada && String(notaCriada).trim() !== '') continue;

    // Cliente: buscar por código se não tem identificador
    if (!row[IdentificadorCliente] || String(row[IdentificadorCliente]).trim() === '') {
      const codigoCliente = row[CodigoCliente];
      if (codigoCliente && typeof codigoCliente !== 'boolean') {
        queries.push({
          sheetRow: i + 1,
          method: 'GET',
          endpoint: `/api/pessoas/codigo/${codigoCliente}`,
          field: 'IdentificadorCliente',
          codigo: codigoCliente
        });
      }
    }

    // Forma de Pagamento
    if (!row[IdentificadorFormaPagamento] || String(row[IdentificadorFormaPagamento]).trim() === '') {
      const codigoFormaPagamento = row[CodigoDaFormaDePagamento];
      if (codigoFormaPagamento && typeof codigoFormaPagamento !== 'boolean') {
        queries.push({
          sheetRow: i + 1,
          method: 'GET',
          endpoint: `/api/formasPagamento`,
          field: 'IdentificadorFormaPagamento',
          codigo: codigoFormaPagamento
        });
      }
    }

    // Operação
    if (!row[IdentificadorOperacao] || String(row[IdentificadorOperacao]).trim() === '') {
      const codigoOperacao = row[CodigoDaOperacao];
      if (codigoOperacao && typeof codigoOperacao !== 'boolean') {
        queries.push({
          sheetRow: i + 1,
          method: 'GET',
          endpoint: `/api/operacoes`,
          field: 'IdentificadorOperacao',
          codigo: codigoOperacao
        });
      }
    }

    // Serviço/Produto
    if (!row[IdentificadorServico] || String(row[IdentificadorServico]).trim() === '') {
      const codigoServico = row[CodigoDoServico];
      if (codigoServico && typeof codigoServico !== 'boolean') {
        queries.push({
          sheetRow: i + 1,
          method: 'GET',
          endpoint: `/api/produtos?codigo=${codigoServico}`,
          field: 'IdentificadorServico',
          codigo: codigoServico
        });
      }
    }

    // Prazo
    if (!row[IdentificadorPrazo] || String(row[IdentificadorPrazo]).trim() === '') {
      const codigoPrazo = row[Codigoprazo];
      if (codigoPrazo && typeof codigoPrazo !== 'boolean') {
        queries.push({
          sheetRow: i + 1,
          method: 'GET',
          endpoint: `/api/prazos`,
          field: 'IdentificadorPrazo',
          codigo: codigoPrazo
        });
      }
    }
  }

  return { 
    queries,
    total: queries.length,
    note: 'Execute cada query no Power Automate e chame applyValidationResults com os resultados'
  };
}

function applyValidationResults(workbook: ExcelScript.Workbook, inputs?: object) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const inputsObj = inputs as { results?: object[] };
  const results = (inputsObj?.results || []);
  let updated = 0;

  for (const resultItem of results) {
    const res = resultItem as { sheetRow?: number; field?: string; value?: string; nome?: string };
    const row = res.sheetRow;
    if (!row || row < 1) continue;

    const field = res.field;
    const value = res.value;

    // Mapear field para índice de coluna
    let colIndex = -1;
    if (field === 'IdentificadorCliente') colIndex = IdentificadorCliente;
    else if (field === 'NomeDoCliente') colIndex = NomeDoCliente;
    else if (field === 'IdentificadorFormaPagamento') colIndex = IdentificadorFormaPagamento;
    else if (field === 'IdentificadorOperacao') colIndex = IdentificadorOperacao;
    else if (field === 'IdentificadorServico') colIndex = IdentificadorServico;
    else if (field === 'IdentificadorPrazo') colIndex = IdentificadorPrazo;

    if (colIndex >= 0 && value) {
      sheet.getRangeByIndexes(row - 1, colIndex, 1, 1).setValue(value);
      updated++;
    }

    // Se veio nome do cliente junto, preencher também
    if (field === 'IdentificadorCliente' && res.nome) {
      sheet.getRangeByIndexes(row - 1, NomeDoCliente, 1, 1).setValue(res.nome);
      updated++;
    }
  }

  return { 
    ok: true, 
    updated,
    message: `${updated} campo(s) atualizado(s) na planilha`
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// SEÇÃO 3: CRIAÇÃO DE PEDIDOS DE VENDA
// ═══════════════════════════════════════════════════════════════════════════

function buildPedidos(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const used = sheet.getUsedRange();
  if (!used) return { payloads: [], total: 0, note: 'Nenhuma área usada na planilha. Sem pedidos a gerar.' };
  const values = used.getValues();

  const payloads: { sheetRow: number; error?: string; payload: object; tipo?: string }[] = [];

  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const notaCriada = row[NotaCriada];

    // Pula linhas que já têm nota criada
    if (notaCriada && String(notaCriada).trim() !== '') continue;

    // Validar campos obrigatórios
    if (!row[IdentificadorCliente] || !row[IdentificadorOperacao] || !row[IdentificadorServico]) {
      payloads.push({
        sheetRow: i + 1,
        error: 'Campos obrigatórios faltando. Execute validação primeiro.',
        payload: {}
      });
      continue;
    }

    // Extrair e formatar valor
    const valorString = String(row[Valor]);
    const valorNumerico = parseValor(valorString);
    const valorFormatado = formatCurrency(valorNumerico);

    // Construir observação detalhada
    const observacao = 
      `${row[NomeDoServico]} (${row[Quantidade]} X ${valorFormatado}) - ${valorNumerico}\n\n\n` +
      `${row[Descriminacao1]}\n\n${row[Descriminacao2]}` +
      `\n\n Data Vencimento: ${row[VencimentoFatura]}`;

    // Montar payload do pedido
    const payload: { [key: string]: unknown } = {
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

    // Adicionar prazo se existir
    if (row[IdentificadorPrazo] && String(row[IdentificadorPrazo]).trim().length > 0) {
      const formaPagamentoEntrada = String(row[FormaPagamentoEntrada]).toUpperCase();
      
      if (formaPagamentoEntrada === 'SIM') {
        payload.Prazo = {
          Identificador: row[IdentificadorPrazo],
          IdentificadorFormaPagamentoEntrada: row[IdentificadorFormaPagamento]
        };
      } else {
        payload.Prazo = {
          Identificador: row[IdentificadorPrazo],
          IdentificadorFormaPagamentoParcelas: row[IdentificadorFormaPagamento]
        };
      }
    }

    payloads.push({ 
      sheetRow: i + 1, 
      payload,
      tipo: 'pedido'
    });
  }

  return { 
    payloads,
    total: payloads.length,
    note: 'POST cada payload para /api/pedidosVenda e chame applyResults com as respostas'
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// SEÇÃO 4: CRIAÇÃO DE DOCUMENTOS FISCAIS
// ═══════════════════════════════════════════════════════════════════════════

function buildDocumentos(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const used = sheet.getUsedRange();
  if (!used) return { payloads: [], total: 0, note: 'Nenhuma área usada na planilha. Sem documentos a gerar.' };
  const values = used.getValues();

  const payloads: { sheetRow: number; error?: string; payload: object; tipo?: string }[] = [];

  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const notaCriada = row[NotaCriada];

    // Pula linhas que já têm nota criada
    if (notaCriada && String(notaCriada).trim() !== '') continue;

    // Validar campos obrigatórios
    if (!row[IdentificadorCliente] || !row[IdentificadorOperacao] || !row[IdentificadorServico]) {
      payloads.push({
        sheetRow: i + 1,
        error: 'Campos obrigatórios faltando. Execute validação primeiro.',
        payload: {}
      });
      continue;
    }

    // Extrair e formatar valor
    const valorString = String(row[Valor]);
    const valorNumerico = parseValor(valorString);
    const valorFormatado = formatCurrency(valorNumerico);

    // Construir observação
    const observacao = 
      `${row[NomeDoServico]} (${row[Quantidade]} X ${valorFormatado}) - ${valorNumerico}\n\n\n` +
      `${row[Descriminacao1]}\n\n${row[Descriminacao2]}` +
      `\n\n Data Vencimento: ${row[VencimentoFatura]}`;

    // Montar payload do documento
    const payload: { [key: string]: unknown } = {
      StatusNotaFiscalEletronica: 'A',
      TipoDocumento: 'S',
      TipoPagamento: '0',
      CodigoEmpresa: row[CodigoDaEmpresa],
      DataEmissao: row[DataEmissao],
      DataReferencia: row[DataEmissao],
      DataReferenciaPagamento: row[DataEmissao],
      IdentificadorOperacao: row[IdentificadorOperacao],
      IdentificadorPessoa: row[IdentificadorCliente],
      Observacao: observacao,
      Itens: [
        {
          CFOP: row[CFOP],
          IdentificadorProduto: row[IdentificadorServico],
          Quantidade: row[Quantidade],
          ValorUnitario: valorNumerico
        }
      ],
      Pagamentos: [
        {
          Aliquota: 100,
          AliquotaConvenio: 10,
          DataVencimento: row[DataEmissao],
          IdentificadorFormaPagamento: row[IdentificadorFormaPagamento],
          Valor: valorNumerico
        }
      ]
    };

    payloads.push({ 
      sheetRow: i + 1, 
      payload,
      tipo: 'documento'
    });
  }

  return { 
    payloads,
    total: payloads.length,
    note: 'POST cada payload para /api/documentos e chame applyResults com as respostas'
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// SEÇÃO 5: APLICAR RESULTADOS
// ═══════════════════════════════════════════════════════════════════════════

function applyResults(workbook: ExcelScript.Workbook, inputs?: object) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const inputsObj = inputs as { results?: object[] };
  const results = (inputsObj?.results || []);
  let updated = 0;

  for (const resultItem of results) {
    const res = resultItem as { sheetRow?: number; notaCriada?: string; retorno?: unknown };
    const row = res.sheetRow;
    if (!row || row < 1) continue;

    // Atualizar NotaCriada
    if (res.notaCriada !== undefined) {
      sheet.getRangeByIndexes(row - 1, NotaCriada, 1, 1).setValue(res.notaCriada);
      updated++;
    }

    // Atualizar RetornoAPI
    if (res.retorno !== undefined) {
      const retornoText = typeof res.retorno === 'string' 
        ? res.retorno 
        : JSON.stringify(res.retorno);
      sheet.getRangeByIndexes(row - 1, RetornoAPI, 1, 1).setValue(retornoText);
      updated++;
    }
  }

  return { 
    ok: true, 
    updated,
    message: `${updated} resultado(s) aplicado(s) na planilha`
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// FUNÇÕES AUXILIARES
// ═══════════════════════════════════════════════════════════════════════════

function parseValor(valorString: string): number {
  if (typeof valorString === 'number') return valorString;
  
  // Remove todos os caracteres exceto dígitos, vírgula e ponto
  const cleaned = String(valorString).replace(/[^\d,.-]/g, '');
  
  // Se tiver vírgula, assume formato BR (1.500,00)
  if (cleaned.includes(',')) {
    return parseFloat(cleaned.replace(/\./g, '').replace(',', '.')) || 0;
  }
  
  // Caso contrário, assume formato US (1500.00)
  return parseFloat(cleaned) || 0;
}

function formatCurrency(value: number): string {
  return new Intl.NumberFormat('pt-BR', {
    style: 'currency',
    currency: 'BRL'
  }).format(value);
}

// ═══════════════════════════════════════════════════════════════════════════
// IMPLEMENTAÇÃO MD5
// ═══════════════════════════════════════════════════════════════════════════

function md5(str: string): string {
  function rotateLeft(lValue: number, iShiftBits: number) {
    return (lValue << iShiftBits) | (lValue >>> (32 - iShiftBits));
  }

  function addUnsigned(lX: number, lY: number) {
    const lX4 = lX & 0x40000000;
    const lY4 = lY & 0x40000000;
    const lX8 = lX & 0x80000000;
    const lY8 = lY & 0x80000000;
    const lResult = (lX & 0x3fffffff) + (lY & 0x3fffffff);
    if (lX4 & lY4) return lResult ^ 0x80000000 ^ lX8 ^ lY8;
    if (lX4 | lY4) {
      if (lResult & 0x40000000) return lResult ^ 0xc0000000 ^ lX8 ^ lY8;
      else return lResult ^ 0x40000000 ^ lX8 ^ lY8;
    } else {
      return lResult ^ lX8 ^ lY8;
    }
  }

  function F(x: number, y: number, z: number) { return (x & y) | (~x & z); }
  function G(x: number, y: number, z: number) { return (x & z) | (y & ~z); }
  function H(x: number, y: number, z: number) { return x ^ y ^ z; }
  function I(x: number, y: number, z: number) { return y ^ (x | ~z); }

  function FF(a: number, b: number, c: number, d: number, x: number, s: number, ac: number) {
    a = addUnsigned(a, addUnsigned(addUnsigned(F(b, c, d), x), ac));
    return addUnsigned(rotateLeft(a, s), b);
  }

  function GG(a: number, b: number, c: number, d: number, x: number, s: number, ac: number) {
    a = addUnsigned(a, addUnsigned(addUnsigned(G(b, c, d), x), ac));
    return addUnsigned(rotateLeft(a, s), b);
  }

  function HH(a: number, b: number, c: number, d: number, x: number, s: number, ac: number) {
    a = addUnsigned(a, addUnsigned(addUnsigned(H(b, c, d), x), ac));
    return addUnsigned(rotateLeft(a, s), b);
  }

  function II(a: number, b: number, c: number, d: number, x: number, s: number, ac: number) {
    a = addUnsigned(a, addUnsigned(addUnsigned(I(b, c, d), x), ac));
    return addUnsigned(rotateLeft(a, s), b);
  }

  function convertToWordArray(strInput: string) {
    const lWordCount: number[] = [];
    let lMessageLength = strInput.length;
    let lNumberOfWordsTempOne = lMessageLength + 8;
    let lNumberOfWordsTempTwo = (lNumberOfWordsTempOne - (lNumberOfWordsTempOne % 64)) / 64;
    let lNumberOfWords = (lNumberOfWordsTempTwo + 1) * 16;
    for (let i = 0; i < lNumberOfWords; i++) lWordCount[i] = 0;
    let lBytePosition = 0;
    let lByteCount = 0;
    while (lByteCount < lMessageLength) {
      const j = (lByteCount - (lByteCount % 4)) / 4;
      lBytePosition = (lByteCount % 4) * 8;
      lWordCount[j] = lWordCount[j] | (strInput.charCodeAt(lByteCount) << lBytePosition);
      lByteCount++;
    }
    const j = (lByteCount - (lByteCount % 4)) / 4;
    lBytePosition = (lByteCount % 4) * 8;
    lWordCount[j] = lWordCount[j] | (0x80 << lBytePosition);
    lWordCount[lNumberOfWords - 2] = lMessageLength << 3;
    lWordCount[lNumberOfWords - 1] = lMessageLength >>> 29;
    return lWordCount;
  }

  function wordToHex(lValue: number) {
    let wordToHexValue = '', wordToHexValueTemp = '', lByte: number, lCount: number;
    for (lCount = 0; lCount <= 3; lCount++) {
      lByte = (lValue >>> (lCount * 8)) & 255;
      wordToHexValueTemp = '0' + lByte.toString(16);
      wordToHexValue = wordToHexValue + wordToHexValueTemp.substr(wordToHexValueTemp.length - 2, 2);
    }
    return wordToHexValue;
  }

  function utf8Encode(str: string) {
    return unescape(encodeURIComponent(str));
  }

  let x = convertToWordArray(utf8Encode(str));
  let a = 0x67452301;
  let b = 0xefcdab89;
  let c = 0x98badcfe;
  let d = 0x10325476;

  for (let k = 0; k < x.length; k += 16) {
    const AA = a, BB = b, CC = c, DD = d;
    a = FF(a, b, c, d, x[k + 0], 7, 0xd76aa478);
    d = FF(d, a, b, c, x[k + 1], 12, 0xe8c7b756);
    c = FF(c, d, a, b, x[k + 2], 17, 0x242070db);
    b = FF(b, c, d, a, x[k + 3], 22, 0xc1bdceee);
    a = FF(a, b, c, d, x[k + 4], 7, 0xf57c0faf);
    d = FF(d, a, b, c, x[k + 5], 12, 0x4787c62a);
    c = FF(c, d, a, b, x[k + 6], 17, 0xa8304613);
    b = FF(b, c, d, a, x[k + 7], 22, 0xfd469501);
    a = FF(a, b, c, d, x[k + 8], 7, 0x698098d8);
    d = FF(d, a, b, c, x[k + 9], 12, 0x8b44f7af);
    c = FF(c, d, a, b, x[k + 10], 17, 0xffff5bb1);
    b = FF(b, c, d, a, x[k + 11], 22, 0x895cd7be);
    a = FF(a, b, c, d, x[k + 12], 7, 0x6b901122);
    d = FF(d, a, b, c, x[k + 13], 12, 0xfd987193);
    c = FF(c, d, a, b, x[k + 14], 17, 0xa679438e);
    b = FF(b, c, d, a, x[k + 15], 22, 0x49b40821);
    a = GG(a, b, c, d, x[k + 1], 5, 0xf61e2562);
    d = GG(d, a, b, c, x[k + 6], 9, 0xc040b340);
    c = GG(c, d, a, b, x[k + 11], 14, 0x265e5a51);
    b = GG(b, c, d, a, x[k + 0], 20, 0xe9b6c7aa);
    a = GG(a, b, c, d, x[k + 5], 5, 0xd62f105d);
    d = GG(d, a, b, c, x[k + 10], 9, 0x02441453);
    c = GG(c, d, a, b, x[k + 15], 14, 0xd8a1e681);
    b = GG(b, c, d, a, x[k + 4], 20, 0xe7d3fbc8);
    a = GG(a, b, c, d, x[k + 9], 5, 0x21e1cde6);
    d = GG(d, a, b, c, x[k + 14], 9, 0xc33707d6);
    c = GG(c, d, a, b, x[k + 3], 14, 0xf4d50d87);
    b = GG(b, c, d, a, x[k + 8], 20, 0x455a14ed);
    a = GG(a, b, c, d, x[k + 13], 5, 0xa9e3e905);
    d = GG(d, a, b, c, x[k + 2], 9, 0xfcefa3f8);
    c = GG(c, d, a, b, x[k + 7], 14, 0x676f02d9);
    b = GG(b, c, d, a, x[k + 12], 20, 0x8d2a4c8a);
    a = HH(a, b, c, d, x[k + 5], 4, 0xfffa3942);
    d = HH(d, a, b, c, x[k + 8], 11, 0x8771f681);
    c = HH(c, d, a, b, x[k + 11], 16, 0x6d9d6122);
    b = HH(b, c, d, a, x[k + 14], 23, 0xfde5380c);
    a = HH(a, b, c, d, x[k + 1], 4, 0xa4beea44);
    d = HH(d, a, b, c, x[k + 4], 11, 0x4bdecfa9);
    c = HH(c, d, a, b, x[k + 7], 16, 0xf6bb4b60);
    b = HH(b, c, d, a, x[k + 10], 23, 0xbebfbc70);
    a = HH(a, b, c, d, x[k + 13], 4, 0x289b7ec6);
    d = HH(d, a, b, c, x[k + 0], 11, 0xeaa127fa);
    c = HH(c, d, a, b, x[k + 3], 16, 0xd4ef3085);
    b = HH(b, c, d, a, x[k + 6], 23, 0x04881d05);
    a = HH(a, b, c, d, x[k + 9], 4, 0xd9d4d039);
    d = HH(d, a, b, c, x[k + 12], 11, 0xe6db99e5);
    c = HH(c, d, a, b, x[k + 15], 16, 0x1fa27cf8);
    b = HH(b, c, d, a, x[k + 2], 23, 0xc4ac5665);
    a = II(a, b, c, d, x[k + 0], 6, 0xf4292244);
    d = II(d, a, b, c, x[k + 7], 10, 0x432aff97);
    c = II(c, d, a, b, x[k + 14], 15, 0xab9423a7);
    b = II(b, c, d, a, x[k + 5], 21, 0xfc93a039);
    a = II(a, b, c, d, x[k + 12], 6, 0x655b59c3);
    d = II(d, a, b, c, x[k + 3], 10, 0x8f0ccc92);
    c = II(c, d, a, b, x[k + 10], 15, 0xffeff47d);
    b = II(b, c, d, a, x[k + 1], 21, 0x85845dd1);
    a = II(a, b, c, d, x[k + 8], 6, 0x6fa87e4f);
    d = II(d, a, b, c, x[k + 15], 10, 0xfe2ce6e0);
    c = II(c, d, a, b, x[k + 6], 15, 0xa3014314);
    b = II(b, c, d, a, x[k + 13], 21, 0x4e0811a1);
    a = II(a, b, c, d, x[k + 4], 6, 0xf7537e82);
    d = II(d, a, b, c, x[k + 11], 10, 0xbd3af235);
    c = II(c, d, a, b, x[k + 2], 15, 0x2ad7d2bb);
    b = II(b, c, d, a, x[k + 9], 21, 0xeb86d391);
    a = addUnsigned(a, AA);
    b = addUnsigned(b, BB);
    c = addUnsigned(c, CC);
    d = addUnsigned(d, DD);
  }

  return (wordToHex(a) + wordToHex(b) + wordToHex(c) + wordToHex(d)).toLowerCase();
}
