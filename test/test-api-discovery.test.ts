/**
 * Descobrir endpoints disponíveis na API Bimer
 */

import { RealAPIClient } from '../src/api/real-api-client';

async function discoverAPI() {
  console.log('🔍 DESCOBRINDO API BIMER\n');
  
  const apiClient = new RealAPIClient();
  
  // Autenticar
  const authResult = await apiClient.authenticate();
  if (!authResult.success) {
    console.log('❌ Falha na autenticação');
    return;
  }
  
  console.log('✅ Autenticado\n');
  
  // Testar endpoints de descoberta
  const discoveryEndpoints = [
    '/',
    '/api',
    '/help',
    '/swagger',
    '/docs',
    '/metadata',
    '/api/help',
    '/api/swagger',
    '/api/docs'
  ];
  
  console.log('🔍 Testando endpoints de descoberta...\n');
  
  for (const endpoint of discoveryEndpoints) {
    console.log(`--- GET ${endpoint} ---`);
    
    try {
      const response = await fetch(`https://087344bimerapi.alterdata.cloud${endpoint}`, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
          'Authorization': `Bearer ${authResult.token}`
        }
      });
      
      console.log(`Status: ${response.status}`);
      
      if (response.ok) {
        const text = await response.text();
        console.log(`✅ ENCONTRADO! Conteúdo (primeiros 200 chars):`);
        console.log(text.substring(0, 200) + '...');
        
        // Se for JSON, tentar fazer parse
        try {
          const json = JSON.parse(text);
          console.log('📋 JSON estruturado:', JSON.stringify(json, null, 2));
        } catch {
          // Não é JSON, mostrar como texto
        }
      } else {
        console.log(`❌ ${response.status}: ${response.statusText}`);
      }
      
    } catch (error) {
      console.log(`💥 Erro: ${error}`);
    }
    
    console.log('');
    await new Promise(resolve => setTimeout(resolve, 300));
  }
  
  // Testar alguns endpoints comuns de APIs REST
  console.log('🔍 Testando endpoints comuns...\n');
  
  const commonEndpoints = [
    '/api/documento',
    '/api/documentos', 
    '/documento',
    '/documentos',
    '/api/nfe',
    '/nfe',
    '/api/nota',
    '/nota'
  ];
  
  for (const endpoint of commonEndpoints) {
    console.log(`--- POST ${endpoint} ---`);
    
    try {
      const response = await fetch(`https://087344bimerapi.alterdata.cloud${endpoint}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Authorization': `Bearer ${authResult.token}`
        },
        body: JSON.stringify({ test: true })
      });
      
      console.log(`Status: ${response.status}`);
      
      if (response.status !== 404) {
        const text = await response.text();
        console.log(`📄 Resposta: ${text.substring(0, 100)}...`);
      }
      
    } catch (error) {
      console.log(`💥 Erro: ${error}`);
    }
    
    await new Promise(resolve => setTimeout(resolve, 300));
  }
}

discoverAPI().catch(console.error);