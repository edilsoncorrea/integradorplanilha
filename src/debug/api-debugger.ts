/**
 * Debugger para pontos de consumo de API
 */

export class APIDebugger {
  private breakpoints: Map<string, boolean> = new Map();
  
  constructor() {
    // Ativa breakpoints por padrão
    this.breakpoints.set('BEFORE_API_CALL', true);
    this.breakpoints.set('AFTER_API_RESPONSE', true);
    this.breakpoints.set('API_ERROR', true);
    this.breakpoints.set('AUTH_CHECK', true);
    this.breakpoints.set('NETWORK_TRACE', true);
  }

  // Breakpoint antes da chamada da API
  beforeAPICall(endpoint: string, payload: any, context?: string, headers?: any) {
    if (this.breakpoints.get('BEFORE_API_CALL')) {
      console.log('\n🔴 BREAKPOINT: Antes da chamada da API');
      console.log('=' .repeat(50));
      console.log(`📍 Contexto: ${context || 'N/A'}`);
      console.log(`🌐 Endpoint: ${endpoint}`);
      console.log(`🕐 Timestamp: ${new Date().toISOString()}`);
      
      if (headers) {
        console.log('🔑 Headers:');
        console.log(JSON.stringify(headers, null, 2));
      }
      
      console.log('📤 Payload a ser enviado:');
      console.log(JSON.stringify(payload, null, 2));
      console.log('=' .repeat(50));
      
      // Simula pausa para debug
      this.pauseForDebug('Pressione Enter para continuar com a chamada da API...');
    }
  }

  // Breakpoint após resposta da API
  afterAPIResponse(endpoint: string, response: any, context?: string) {
    if (this.breakpoints.get('AFTER_API_RESPONSE')) {
      console.log('\n🟢 BREAKPOINT: Após resposta da API');
      console.log('=' .repeat(50));
      console.log(`📍 Contexto: ${context || 'N/A'}`);
      console.log(`🌐 Endpoint: ${endpoint}`);
      console.log('📥 Resposta recebida:');
      console.log(JSON.stringify(response, null, 2));
      console.log('=' .repeat(50));
      
      this.pauseForDebug('Pressione Enter para continuar...');
    }
  }

  // Breakpoint em caso de erro
  onAPIError(endpoint: string, error: any, context?: string) {
    if (this.breakpoints.get('API_ERROR')) {
      console.log('\n🔴 BREAKPOINT: Erro na API');
      console.log('=' .repeat(50));
      console.log(`📍 Contexto: ${context || 'N/A'}`);
      console.log(`🌐 Endpoint: ${endpoint}`);
      console.log('❌ Erro:');
      console.log(error);
      console.log('=' .repeat(50));
      
      this.pauseForDebug('Pressione Enter para continuar...');
    }
  }

  // Simula pausa para debug (em ambiente real seria um debugger)
  private pauseForDebug(message: string) {
    console.log(`\n⏸️  ${message}`);
    console.log('💡 Em ambiente real, use: debugger; ou breakpoint do IDE');
    
    // Em Node.js, simula pausa
    if (typeof process !== 'undefined') {
      console.log('🔍 Dados disponíveis para inspeção no contexto atual');
    }
  }

  // Controle de breakpoints
  enableBreakpoint(type: string) {
    this.breakpoints.set(type, true);
    console.log(`✅ Breakpoint ${type} ativado`);
  }

  disableBreakpoint(type: string) {
    this.breakpoints.set(type, false);
    console.log(`❌ Breakpoint ${type} desativado`);
  }

  // Breakpoint para verificação de autenticação
  checkAuth(token?: string, context?: string) {
    if (this.breakpoints.get('AUTH_CHECK')) {
      console.log('\n🔐 BREAKPOINT: Verificação de Autenticação');
      console.log('=' .repeat(50));
      console.log(`📍 Contexto: ${context || 'N/A'}`);
      console.log(`🔑 Token presente: ${token ? 'SIM' : 'NÃO'}`);
      if (token) {
        console.log(`🔑 Token (primeiros 20 chars): ${token.substring(0, 20)}...`);
      }
      console.log('=' .repeat(50));
      
      this.pauseForDebug('Verificar autenticação...');
    }
  }

  // Trace de rede
  networkTrace(method: string, url: string, status?: number, timing?: number) {
    if (this.breakpoints.get('NETWORK_TRACE')) {
      console.log('\n🌐 NETWORK TRACE');
      console.log(`${method.toUpperCase()} ${url}`);
      if (status) console.log(`Status: ${status}`);
      if (timing) console.log(`Tempo: ${timing}ms`);
      console.log('');
    }
  }

  listBreakpoints() {
    console.log('\n📋 Status dos Breakpoints:');
    this.breakpoints.forEach((enabled, type) => {
      console.log(`   ${enabled ? '✅' : '❌'} ${type}`);
    });
  }
}

// Instância global do debugger
export const apiDebugger = new APIDebugger();