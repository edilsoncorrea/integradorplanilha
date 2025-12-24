/**
 * Teste com diferentes formatos de payload
 */

import { RealAPIClient } from '../src/api/real-api-client';

async function testDifferentFormats() {
  console.log('🔧 TESTE DE DIFERENTES FORMATOS\n');
  
  const apiClient = new RealAPIClient();
  
  // Override para testar form-encoded
  const originalCallAPI = apiClient.callAPI;
  apiClient.callAPI = async function(endpoint: string, payload: any, context?: string) {
    const fullURL = `https://087344bimerapi.alterdata.cloud${endpoint}`;
    
    if (endpoint.includes('/oauth/token')) {
      console.log('🔧 Testando formato form-encoded...');
      
      // Formato form-encoded
      const formData = new URLSearchParams();
      Object.keys(payload).forEach(key => {
        formData.append(key, payload[key]);
      });
      
      try {
        const response = await fetch(fullURL, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept': 'application/json'
          },
          body: formData.toString()
        });
        
        console.log(`Status: ${response.status}`);
        const data = await response.text();
        console.log(`Resposta: ${data}`);
        
        if (response.ok) {
          return { success: true, data: JSON.parse(data) };
        } else {
          return { success: false, error: data };
        }
      } catch (error) {
        return { success: false, error };
      }
    }
    
    return originalCallAPI.call(this, endpoint, payload, context);
  };
  
  const result = await apiClient.authenticate();
  console.log('\n📊 Resultado:', result);
}

testDifferentFormats().catch(console.error);