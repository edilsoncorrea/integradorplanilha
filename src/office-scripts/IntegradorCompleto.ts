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
 * 
 * IMPORTANTE:
 * ────────────────────────────────────────────────────────────────────────────
 * Este script NÃO faz chamadas HTTP diretamente. As chamadas à API devem ser
 * feitas pelo Power Automate, que:
 *   1. Chama este script para gerar payloads
 *   2. Executa as requisições HTTP na API Bimer
 *   3. Chama este script novamente para aplicar os resultados na planilha
 * 
 * USO:
 * ────────────────────────────────────────────────────────────────────────────
 * Use o parâmetro 'action' em inputs para escolher a operação:
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
 * EXEMPLO DE USO NO POWER AUTOMATE:
 * ────────────────────────────────────────────────────────────────────────────
 * 
 * Passo 1: Autenticar
 *   Script: action = 'buildAuthPayload'
 *   HTTP POST: Usa URL e payload retornados
 *   Salva: access_token para próximas chamadas
 * 
 * Passo 2: Validar planilha
 *   Script: action = 'buildValidationQueries'
 *   HTTP GET: Para cada query retornada
 *   Script: action = 'applyValidationResults' com resultados
 * 
 * Passo 3: Criar pedidos/documentos
 *   Script: action = 'buildPedidos' OU 'buildDocumentos'
 *   HTTP POST: Para cada payload retornado
 *   Script: action = 'applyResults' com respostas da API
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

// ═══════════════════════════════════════════════════════════════════════════
// PONTO DE ENTRADA PRINCIPAL
// ═══════════════════════════════════════════════════════════════════════════

async function main(workbook: ExcelScript.Workbook, inputs?: { action?: string; host?: string; username?: string; senha?: string; nonce?: string; value?: string; results?: object[]; executeAPI?: boolean }): Promise<object> {
  const action = inputs?.action || 'criarDocumentosCompleto';
  const executeAPI = inputs?.executeAPI !== false; // Default true

  // FLUXO COMPLETO (NOVO - COM EXECUÇÃO AUTOMÁTICA)
  if (action === 'criarDocumentosCompleto') {
    if (executeAPI) {
      return await executarFluxoCompleto(workbook, inputs);
    } else {
      return buildDocumentos(workbook);
    }
  }

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
    return {
      message: 'INTEGRADOR COMPLETO - Office Scripts para Excel Online',
      actions: {
        autenticacao: ['buildAuthPayload', 'hash'],
        validacao: ['buildValidationQueries', 'applyValidationResults'],
        pedidos: ['buildPedidos'],
        documentos: ['buildDocumentos'],
        resultados: ['applyResults'],
        completo: ['criarDocumentosCompleto']
      },
      usage: 'Passe inputs.action com uma das ações listadas acima',
      exemplo: '{ action: "buildPedidos" }'
    };
  }

  return { error: `Ação desconhecida: ${action}` };
}

// ═══════════════════════════════════════════════════════════════════════════
// SEÇÃO 0: FLUXO COMPLETO COM FETCH (NOVO)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Executa o fluxo completo: autenticação → criar documentos → atualizar planilha
 * USA FETCH para fazer chamadas HTTP diretas à API Bimer
 * 
 * @param workbook Workbook do Excel
 * @param inputs Configurações opcionais (host, username, senha)
 * @returns Resultado da execução completa
 */
async function executarFluxoCompleto(workbook: ExcelScript.Workbook, inputs?: { host?: string; username?: string; senha?: string; nonce?: string }): Promise<object> {
  const host = (inputs?.host as string) || HOST;
  const username = (inputs?.username as string) || 'supervisor';
  const senha = (inputs?.senha as string) || 'Senhas123';
  const nonce = (inputs?.nonce as string) || '123456789';

  try {
    // ═══════════════════════════════════════════════════════════════════════
    // PASSO 1: AUTENTICAÇÃO
    // ═══════════════════════════════════════════════════════════════════════
    console.log('🔐 Autenticando na API Bimer...');
    
    const password = md5(username + nonce + senha);
    const authUrl = host.endsWith('/') ? host + 'oauth/token' : host + '/oauth/token';
    
    const authBody = new URLSearchParams({
      client_id: 'IntegracaoBimer.js',
      username: username,
      password: password,
      grant_type: 'password',
      nonce: nonce
    });

    const authResponse = await fetch(authUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: authBody.toString()
    });

    if (!authResponse.ok) {
      return { 
        error: 'Falha na autenticação', 
        status: authResponse.status,
        statusText: authResponse.statusText
      };
    }

    const authJson = await authResponse.json() as { access_token?: string };
    const token = authJson.access_token;

    if (!token) {
      return { error: 'Token não retornado pela API' };
    }

    console.log('✅ Autenticado com sucesso!');

    // ═══════════════════════════════════════════════════════════════════════
    // PASSO 2: GERAR PAYLOADS DOS DOCUMENTOS
    // ═══════════════════════════════════════════════════════════════════════
    console.log('📄 Gerando payloads dos documentos...');
    
    const buildResult = buildDocumentos(workbook);
    
    if ('error' in buildResult) {
      return buildResult;
    }

    const payloads = buildResult.payloads || [];
    console.log(`📦 ${payloads.length} documento(s) para processar`);

    if (payloads.length === 0) {
      return { message: 'Nenhum documento para processar', total: 0 };
    }

    // ═══════════════════════════════════════════════════════════════════════
    // PASSO 3: CRIAR DOCUMENTOS NA API E COLETAR RESULTADOS
    // ═══════════════════════════════════════════════════════════════════════
    console.log('🚀 Criando documentos na API...');
    
    const apiUrl = host.endsWith('/') ? host + 'api/documentos' : host + '/api/documentos';
    const results: { sheetRow: number; notaCriada: string; retorno: string }[] = [];
    let sucessos = 0;
    let erros = 0;

    for (const item of payloads) {
      const sheetRow = item.sheetRow;
      
      // Pula linhas com erro de validação
      if (item.error) {
        results.push({
          sheetRow: sheetRow,
          notaCriada: 'Não',
          retorno: item.error
        });
        erros++;
        continue;
      }

      try {
        const response = await fetch(apiUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
          },
          body: JSON.stringify(item.payload)
        });

        if (response.ok) {
          const responseJson = await response.json() as { Identificador?: string };
          const identificador = responseJson.Identificador || 'OK';
          
          results.push({
            sheetRow: sheetRow,
            notaCriada: 'Sim',
            retorno: `Documento criado: ${identificador}`
          });
          
          sucessos++;
          console.log(`✅ Linha ${sheetRow}: ${identificador}`);
        } else {
          const errorText = await response.text();
          results.push({
            sheetRow: sheetRow,
            notaCriada: 'Não',
            retorno: `Erro ${response.status}: ${errorText}`
          });
          
          erros++;
          console.log(`❌ Linha ${sheetRow}: ${response.status}`);
        }
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : String(err);
        results.push({
          sheetRow: sheetRow,
          notaCriada: 'Não',
          retorno: `Exceção: ${errorMsg}`
        });
        
        erros++;
        console.log(`❌ Linha ${sheetRow}: ${errorMsg}`);
      }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // PASSO 4: ATUALIZAR PLANILHA COM RESULTADOS
    // ═══════════════════════════════════════════════════════════════════════
    console.log('📝 Atualizando planilha com resultados...');
    
    const applyResult = applyResults(workbook, { results });

    return {
      ok: true,
      processados: payloads.length,
      sucessos: sucessos,
      erros: erros,
      planilhaAtualizada: applyResult.updated,
      message: `✅ ${sucessos} documento(s) criado(s) | ❌ ${erros} erro(s)`
    };

  } catch (err) {
    const errorMsg = err instanceof Error ? err.message : String(err);
    console.error('❌ Erro no fluxo completo:', errorMsg);
    return { 
      error: 'Erro no fluxo completo', 
      detalhes: errorMsg 
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SEÇÃO 1: AUTENTICAÇÃO
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Gera o payload de autenticação para a API Bimer
 * 
 * @param inputs Deve conter: host (opcional), username, senha, nonce (opcional)
 * @returns { url, method, payload } para ser usado no Power Automate
 */
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

/**
 * Calcula hash MD5 de um valor
 * 
 * @param inputs Deve conter: value (string a ser hasheada)
 * @returns { md5: string }
 */
function hashValue(inputs?: { value?: string }) {
  const value = (inputs?.value as string) || '';
  return { md5: md5(value) };
}

// ═══════════════════════════════════════════════════════════════════════════
// SEÇÃO 2: VALIDAÇÃO DA PLANILHA
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Analisa a planilha e retorna lista de consultas GET necessárias para
 * buscar identificadores faltantes (Cliente, Operação, Serviço, etc.)
 * 
 * @param workbook Workbook do Excel
 * @returns { queries: Array<{sheetRow, method, endpoint, field}> }
 */
function buildValidationQueries(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const used = sheet.getUsedRange();
  if (!used) return { queries: [] };
  const values = used.getValues();

  const queries: { sheetRow: number; method: string; endpoint: string; field: string; codigo: string | number }[] = [];

  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const notaCriada = row[NotaCriada];
    
    // Pula linhas que já têm nota criada
    if (notaCriada && String(notaCriada).trim() !== '') continue;

    // Cliente: buscar por código se não tem identificador
    if (!row[IdentificadorCliente] || String(row[IdentificadorCliente]).trim() === '') {
      const codigoCliente = row[CodigoCliente];
      if (codigoCliente) {
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
      if (codigoFormaPagamento) {
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
      if (codigoOperacao) {
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
      if (codigoServico) {
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
      if (codigoPrazo) {
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

/**
 * Aplica os resultados das consultas de validação na planilha
 * 
 * @param workbook Workbook do Excel
 * @param inputs Deve conter: results = Array<{sheetRow, field, value, nome?}>
 * @returns { ok: boolean, updated: number }
 */
function applyValidationResults(workbook: ExcelScript.Workbook, inputs?: { results?: object[] }) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const results = (inputs?.results as { sheetRow: number; field: string; value: string; nome?: string }[]) || [];
  let updated = 0;

  for (const res of results) {
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

/**
 * Gera payloads de pedidos de venda a partir das linhas da planilha
 * 
 * @param workbook Workbook do Excel
 * @returns { payloads: Array<{sheetRow, payload}> }
 */
function buildPedidos(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const used = sheet.getUsedRange();
  if (!used) return { payloads: [] };
  const values = used.getValues();

  const payloads: { sheetRow: number; error?: string; payload: unknown; tipo?: string }[] = [];

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

/**
 * Gera payloads de documentos fiscais a partir das linhas da planilha
 * 
 * @param workbook Workbook do Excel
 * @returns { payloads: Array<{sheetRow, payload}> }
 */
function buildDocumentos(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const used = sheet.getUsedRange();
  if (!used) return { payloads: [] };
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

/**
 * Aplica resultados das chamadas da API de volta na planilha
 * 
 * @param workbook Workbook do Excel
 * @param inputs Deve conter: results = Array<{sheetRow, notaCriada?, retorno?}>
 * @returns { ok: boolean, updated: number }
 */
function applyResults(workbook: ExcelScript.Workbook, inputs?: { results?: object[] }) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return { error: 'Planilha "Documento" não encontrada' };

  const results = (inputs?.results as { sheetRow: number; notaCriada?: string; retorno?: string }[]) || [];
  let updated = 0;

  for (const res of results) {
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

/**
 * Extrai valor numérico de string formatada (ex: "R$ 1.500,00" -> 1500)
 */
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

/**
 * Formata número como moeda brasileira
 */
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
