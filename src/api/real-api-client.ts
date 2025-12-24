/**
 * Cliente real para chamadas de API com debug completo
 */

import { apiDebugger } from '../debug/api-debugger';

export class RealAPIClient {
  private baseURL: string;
  private token?: string;

  //ATENCAO EDILSON
  /*
  private credentials = {
    host: 'https://087344bimerapi.alterdata.cloud',
    username: 'bimerapi',  // Do Autenticacao.ts
    senha: '123456',       // Do Autenticacao.ts
    nonce: '123456789'     // Do Autenticacao.ts
  };
  */

  private credentials = {
    host: 'https://homologacaowisepcp.alterdata.com.br/BimerApi',
    username: 'supervisor',  // Do Autenticacao.ts
    senha: 'Senhas123',       // Do Autenticacao.ts
    nonce: '123456789'     // Do Autenticacao.ts
  };

  constructor(baseURL?: string) {
    this.baseURL = baseURL || this.credentials.host;
  }

  // Configurar token de autenticação
  setAuthToken(token: string) {
    this.token = token;
    apiDebugger.checkAuth(token, 'Token configurado');
  }

  // Autenticar com API Bimer usando dados do Autenticacao.ts
  async authenticate() {
    console.log('🔐 Autenticando com API Bimer...');
    
    // Usar função do Autenticacao.ts para gerar payload
    const { main } = await import('../office-scripts/Autenticacao');
    
    const authData = main(null as any, {
      action: 'buildAuthPayload',
      host: this.baseURL,
      username: this.credentials.username,
      senha: this.credentials.senha,
      nonce: this.credentials.nonce
    });
    
    try {
      const result = await this.callAPI('/oauth/token', authData.payload, 'Autenticação Bimer');
      
      if (result.success && result.data) {
        // Usar token da resposta
        const token = (result.data as any)?.access_token || `token_${Date.now()}`;
        this.setAuthToken(token);
        console.log('✅ Autenticação bem-sucedida');
        return { success: true, token };
      } else {
        console.log('❌ Falha na autenticação:', result.error);
        return { success: false, error: result.error || 'Falha na autenticação' };
      }
    } catch (error) {
      console.log('❌ Erro na autenticação:', error);
      return { success: false, error };
    }
  }

  // Fazer chamada real para API
  async callAPI(endpoint: string, payload: any, context?: string) {
    const fullURL = `${this.baseURL}${endpoint}`;
    const startTime = Date.now();

    // Headers com autenticação
    const headers: any = {
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    };

    if (this.token) {
      headers['Authorization'] = `Bearer ${this.token}`;
    }

    // 🔴 BREAKPOINT: Antes da chamada
    apiDebugger.beforeAPICall(fullURL, payload, context, headers);

    try {
      // 🌐 NETWORK TRACE: Início
      apiDebugger.networkTrace('POST', fullURL);

      // Fazer chamada real
      let body: string;
      let contentType = headers['Content-Type'];
      
      // Para /oauth/token usar form-encoded
      if (endpoint.includes('/oauth/token')) {
        const formData = new URLSearchParams();
        Object.keys(payload).forEach(key => {
          formData.append(key, payload[key]);
        });
        body = formData.toString();
        contentType = 'application/x-www-form-urlencoded';
      } else {
        body = JSON.stringify(payload);
      }
      
      const response = await fetch(fullURL, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': contentType },
        body
      });

      const timing = Date.now() - startTime;
      apiDebugger.networkTrace('POST', fullURL, response.status, timing);

      // Verificar se há conteúdo para fazer parse
      const responseText = await response.text();
      let data: any = null;
      
      if (responseText.trim()) {
        try {
          data = JSON.parse(responseText);
        } catch (parseError) {
          data = responseText; // Usar texto se não for JSON válido
        }
      }
      
      if (response.ok) {
        // 🟢 BREAKPOINT: Sucesso
        apiDebugger.afterAPIResponse(fullURL, data, context);
        return { success: true, data };
      } else {
        // 🔴 BREAKPOINT: Erro HTTP
        apiDebugger.onAPIError(fullURL, { status: response.status, data, responseText }, context);
        return { success: false, error: data || `HTTP ${response.status}: ${response.statusText}` };
      }

    } catch (error) {
      const errorTiming = Date.now() - startTime;
      apiDebugger.networkTrace('POST', fullURL, 0, errorTiming);
      
      // 🔴 BREAKPOINT: Erro de rede
      apiDebugger.onAPIError(fullURL, error, `${context} - Erro de rede`);
      
      return { success: false, error };
    }
  }

  // Criar pedido de venda
  async createPedido(payload: any) {
    console.log('📦 Criando pedido de venda...');
    
    if (!this.token) {
      console.log('❌ Token de autenticação necessário');
      return { success: false, error: 'Token de autenticação necessário' };
    }
    
    try {
      const result = await this.callAPI('/api/venda/pedidos', payload, 'Criar Pedido de Venda');
      
      if (result.success) {
        console.log('✅ Pedido criado com sucesso');
        return result;
      } else {
        console.log('❌ Falha ao criar pedido:', result.error);
        return result;
      }
    } catch (error) {
      console.log('❌ Erro ao criar pedido:', error);
      return { success: false, error };
    }
  }

  // Testar conectividade
  async testConnection() {
    console.log('🔌 Testando conectividade com a API...\n');
    
    try {
      const result = await this.callAPI('/health', {}, 'Teste de conectividade');
      
      if (result.success) {
        console.log('✅ API acessível');
      } else {
        console.log('❌ API não acessível');
      }
      
      return result;
    } catch (error) {
      console.log('❌ Erro de conectividade:', error);
      return { success: false, error };
    }
  }
}