/**
 * Mock completo da API ExcelScript para testes locais
 */

export class MockWorksheet {
  private name: string;
  private data: any[][] = [];
  private usedRange: MockRange | null = null;

  constructor(name: string, initialData?: any[][]) {
    this.name = name;
    if (initialData) {
      this.data = initialData;
      this.usedRange = new MockRange(initialData, 0, 0, initialData.length, initialData[0]?.length || 0, this);
    }
  }

  getName(): string {
    return this.name;
  }

  getUsedRange(): MockRange | null {
    return this.usedRange;
  }

  getRangeByIndexes(row: number, column: number, rowCount: number, columnCount: number): MockRange {
    return new MockRange(this.data, row, column, rowCount, columnCount, this);
  }

  setData(data: any[][]) {
    this.data = data;
    this.usedRange = new MockRange(data, 0, 0, data.length, data[0]?.length || 0, this);
  }

  // Método para atualizar dados internamente
  updateData(newData: any[][]) {
    this.data = newData;
    this.usedRange = new MockRange(newData, 0, 0, newData.length, newData[0]?.length || 0, this);
  }
}

export class MockRange {
  private data: any[][];
  private startRow: number = 0;
  private startColumn: number = 0;
  private rowCount: number;
  private columnCount: number;
  private worksheet: MockWorksheet | null = null;

  constructor(data: any[][], startRow?: number, startColumn?: number, rowCount?: number, columnCount?: number, worksheet?: MockWorksheet) {
    this.data = data;
    this.startRow = startRow || 0;
    this.startColumn = startColumn || 0;
    this.rowCount = rowCount || data.length;
    this.columnCount = columnCount || (data[0]?.length || 0);
    this.worksheet = worksheet || null;
  }

  getValues(): any[][] {
    return this.data;
  }

  setValue(value: any): void {
    if (this.data[this.startRow]) {
      this.data[this.startRow][this.startColumn] = value;
      // Atualizar worksheet se disponível
      if (this.worksheet) {
        this.worksheet.updateData(this.data);
      }
    }
  }
}

export class MockWorkbook {
  private worksheets: Map<string, MockWorksheet> = new Map();
  public realData: any[][] = []; // Vetor público para simulação

  constructor() {
    const sampleData = [
      ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
      ['Cd.empresa', 'Codigo Cliente', 'Nome do Cliente', 'Identificador do Cliente', 'Cd. da Operacao', 'Identificador da Operacao', 'CFOP', 'Cd. do serviço', 'Identificador do Serviço', 'Nome do serviço', 'Quantidade', 'Valor', 'Descriminação 1', 'Descriminação 2', 'Cd prazo', 'Identificador prazo', 'Forma Pagamento é entrada?', 'Código da F. pagamento', 'Identificador da  forma de Pagamento', 'Dt.Emissao', 'Vencimento', 'Nota Criada?', 'Retorno'],
      ['000001', '000146', '3WLINK INTERNET LTDA ME', '00A000005C', '000019', '00A000000R', '6.101', '000896', '00A00000R3', 'ACESSÓRIO E FERRAGEM ADESIVO AC1125', '10', 'R$ 1.000,00', 'Desc 1', 'Desc 2', '000002', '00A000000H', 'Não', '000002', '00A0000003', '04/12/2025', '04/12/2025', '', '']
    ];

    this.worksheets.set('Documento', new MockWorksheet('Documento', sampleData));
  }

  loadRealData() {
    // Tentar carregar do CSV primeiro, senão usar dados padrão
    try {
      const { CSVSimulator } = require('./csv-simulator');
      this.realData = CSVSimulator.carregarCSV();
      console.log('📄 Dados carregados do CSV para simulação');
    } catch (error) {
      // Fallback para dados padrão se CSV não existir
      this.realData = [
        ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['Cd.empresa', 'Codigo Cliente', 'Nome do Cliente', 'Identificador do Cliente', 'Cd. da Operacao', 'Identificador da Operacao', 'CFOP', 'Cd. do serviço', 'Identificador do Serviço', 'Nome do serviço', 'Quantidade', 'Valor', 'Descriminação 1', 'Descriminação 2', 'Cd prazo', 'Identificador prazo', 'Forma Pagamento é entrada?', 'Código da F. pagamento', 'Identificador da  forma de Pagamento', 'Dt.Emissao', 'Vencimento', 'Nota Criada?', 'Retorno'],
        ['000001', '000146', '3WLINK INTERNET LTDA ME', '00A000005C', '000019', '00A000000R', '6.101', '000896', '00A00000R3', 'PRODUTO A', '10', 'R$ 1.000,00', '', '', '000002', '00A000000H', 'Não', '000002', '00A0000003', '05/12/2025', '05/12/2026', '', ''],
        ['000001', '000146', '3WLINK INTERNET LTDA ME', '00A000005C', '000019', '00A000000R', '6.101', '000896', '00A00000R3', 'PRODUTO A', '10', 'R$ 2.000,00', '', '', '000002', '00A000000H', 'Não', '000002', '00A0000003', '05/12/2025', '05/12/2026', '', ''],
        ['000001', '000146', '3WLINK INTERNET LTDA ME', '00A000005C', '000019', '00A000000R', '6.101', '000896', '00A00000R3', 'PRODUTO A', '10', 'R$ 3.000,00', '', '', '000002', '00A000000H', 'Não', '000002', '00A0000003', '05/12/2025', '05/12/2026', '', '']
      ];
      console.log('📄 Usando dados padrão (CSV não encontrado)');
    }

    this.worksheets.set('Documento', new MockWorksheet('Documento', this.realData));
    return this;
  }

  getWorksheet(name: string): MockWorksheet | null {
    return this.worksheets.get(name) || null;
  }

  addWorksheet(name: string, data?: any[][]): MockWorksheet {
    const worksheet = new MockWorksheet(name, data);
    this.worksheets.set(name, worksheet);
    return worksheet;
  }

  // Método para inspecionar o vetor realData
  getRealDataVector(): any[][] {
    return this.realData;
  }

  // Método para mostrar o vetor atualizado como código
  exportRealDataAsCode(): string {
    const formatValue = (val: any) => {
      if (val === '') return "''";
      if (typeof val === 'string') return `'${val}'`;
      return `'${val}'`;
    };

    let code = 'const realDataAtualizado = [\n';
    this.realData.forEach((linha, index) => {
      const linhaFormatada = linha.map(formatValue).join(', ');
      code += `      [${linhaFormatada}]${index < this.realData.length - 1 ? ',' : ''}\n`;
    });
    code += '    ];';
    return code;
  }
}

// Função utilitária para atualizar planilha com resultado da API
export function atualizarPlanilhaComResultado(workbook: any, sheetRow: number, sucesso: boolean, identificador: string) {
  const sheet = workbook.getWorksheet('Documento');
  if (!sheet) return;

  // Importar constantes
  const NotaCriada = 21; // Coluna "Nota Criada?"
  const RetornoAPI = 22;  // Coluna "Retorno"
  
  // Obter dados atuais da planilha
  const usedRange = sheet.getUsedRange();
  if (!usedRange) return;
  
  const data = usedRange.getValues();
  
  // Atualizar os dados diretamente no array
  if (data[sheetRow - 1]) {
    data[sheetRow - 1][NotaCriada] = sucesso ? 'Sim' : 'Não';
    data[sheetRow - 1][RetornoAPI] = identificador;
    
    // Atualizar a planilha com os novos dados
    sheet.updateData(data);
    
    // IMPORTANTE: Atualizar também o vetor realData para simulação
    if (workbook.realData && workbook.realData[sheetRow - 1]) {
      workbook.realData[sheetRow - 1][NotaCriada] = sucesso ? 'Sim' : 'Não';
      workbook.realData[sheetRow - 1][RetornoAPI] = identificador;
    }
    
    // NOVO: Atualizar também o arquivo CSV para persistência real
    const { CSVSimulator } = require('./csv-simulator');
    CSVSimulator.atualizarLinha(sheetRow, sucesso, identificador);
    
    console.log(`📝 Planilha atualizada - Linha ${sheetRow}: NotaCriada="${sucesso ? 'Sim' : 'Não'}", Retorno="${identificador}"`);
  }
}

// Função para mostrar o vetor atualizado
export function mostrarVetorAtualizado(workbook: MockWorkbook) {
  console.log('\n📋 VETOR REALDATA ATUALIZADO:');
  console.log(workbook.exportRealDataAsCode());
}

// Tipos globais para simular ExcelScript
declare global {
  namespace ExcelScript {
    interface Workbook {
      getWorksheet(name: string): Worksheet | null;
    }
    
    interface Worksheet {
      getName(): string;
      getUsedRange(): Range | null;
      getRangeByIndexes(row: number, column: number, rowCount: number, columnCount: number): Range;
    }
    
    interface Range {
      getValues(): any[][];
      setValue(value: any): void;
    }
  }
}