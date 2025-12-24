/**
 * Simulador CSV - Réplica da planilha Excel
 */

import * as fs from 'fs';
import * as path from 'path';

const CSV_PATH = path.join(__dirname, '../../data/planilha-simulacao.csv');

// Garantir que o diretório existe
const dataDir = path.dirname(CSV_PATH);
if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
}

export class CSVSimulator {
  
  // Criar CSV inicial com dados limpos
  static criarCSVInicial(): void {
    const dadosIniciais = [
      ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
      ['Cd.empresa', 'Codigo Cliente', 'Nome do Cliente', 'Identificador do Cliente', 'Cd. da Operacao', 'Identificador da Operacao', 'CFOP', 'Cd. do serviço', 'Identificador do Serviço', 'Nome do serviço', 'Quantidade', 'Valor', 'Descriminação 1', 'Descriminação 2', 'Cd prazo', 'Identificador prazo', 'Forma Pagamento é entrada?', 'Código da F. pagamento', 'Identificador da  forma de Pagamento', 'Dt.Emissao', 'Vencimento', 'Nota Criada?', 'Retorno'],
      ['000001', '000146', '3WLINK INTERNET LTDA ME', '00A000005C', '000019', '00A000000R', '6.101', '000896', '00A00000R3', 'PRODUTO A', '10', 'R$ 1.000,00', '', '', '000002', '00A000000H', 'Não', '000002', '00A0000003', '05/12/2025', '05/12/2026', '', ''],
      ['000001', '000146', '3WLINK INTERNET LTDA ME', '00A000005C', '000019', '00A000000R', '6.101', '000896', '00A00000R3', 'PRODUTO A', '10', 'R$ 2.000,00', '', '', '000002', '00A000000H', 'Não', '000002', '00A0000003', '05/12/2025', '05/12/2026', '', ''],
      ['000001', '000146', '3WLINK INTERNET LTDA ME', '00A000005C', '000019', '00A000000R', '6.101', '000896', '00A00000R3', 'PRODUTO A', '10', 'R$ 3.000,00', '', '', '000002', '00A000000H', 'Não', '000002', '00A0000003', '05/12/2025', '05/12/2026', '', '']
    ];
    
    this.salvarCSV(dadosIniciais);
    console.log(`📄 CSV inicial criado: ${CSV_PATH}`);
  }
  
  // Carregar dados do CSV
  static carregarCSV(): any[][] {
    if (!fs.existsSync(CSV_PATH)) {
      this.criarCSVInicial();
    }
    
    const csvContent = fs.readFileSync(CSV_PATH, 'utf-8');
    const linhas = csvContent.split('\n').filter(linha => linha.trim());
    
    return linhas.map(linha => {
      // Parse usando ponto e vírgula como separador
      const colunas = linha.split(';');
      // Garantir que temos exatamente 23 colunas
      while (colunas.length < 23) {
        colunas.push('');
      }
      return colunas.slice(0, 23).map(valor => valor.trim());
    });
  }
  
  // Salvar dados no CSV
  static salvarCSV(dados: any[][]): void {
    const csvContent = dados.map(linha => {
      // Garantir que cada linha tem exatamente 23 colunas
      const linhaPadronizada = [...linha];
      while (linhaPadronizada.length < 23) {
        linhaPadronizada.push('');
      }
      return linhaPadronizada.slice(0, 23).map(valor => `${valor || ''}`).join(';');
    }).join('\n');
    
    fs.writeFileSync(CSV_PATH, csvContent, 'utf-8');
  }
  
  // Atualizar linha específica no CSV
  static atualizarLinha(numeroLinha: number, sucesso: boolean, identificador: string): void {
    const dados = this.carregarCSV();
    
    if (dados[numeroLinha - 1]) {
      dados[numeroLinha - 1][21] = sucesso ? 'Sim' : 'Não'; // NotaCriada
      dados[numeroLinha - 1][22] = identificador;           // Retorno
      
      this.salvarCSV(dados);
      console.log(`📝 CSV atualizado - Linha ${numeroLinha}: NotaCriada="${sucesso ? 'Sim' : 'Não'}", Retorno="${identificador}"`);
    }
  }
  
  // Resetar CSV para estado inicial
  static resetarCSV(): void {
    this.criarCSVInicial();
    console.log('🔄 CSV resetado para estado inicial');
  }
  
  // Mostrar estado atual do CSV
  static mostrarEstado(): void {
    const dados = this.carregarCSV();
    console.log('\n📊 ESTADO ATUAL DO CSV:');
    
    dados.slice(2).forEach((linha, index) => {
      const numeroLinha = index + 3;
      const notaCriada = linha[21] || '';
      const retorno = linha[22] || '';
      console.log(`  Linha ${numeroLinha}: NotaCriada="${notaCriada}", Retorno="${retorno}"`);
    });
  }
  
  // Obter caminho do arquivo CSV
  static getCaminhoCSV(): string {
    return CSV_PATH;
  }
}