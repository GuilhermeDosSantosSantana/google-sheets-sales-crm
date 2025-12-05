// --- CONFIGURAÇÕES GERAIS ---
const NOME_ABA_PROSPECCAO = "Prospecção";
const NOME_ABA_CLIENTES = "Clientes";
const NOME_ABA_PERDIDOS = "Perdidos";

// MAPEAMENTO DE COLUNAS (REALINHADO: Categoria na F, o resto andou +1)
// A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12, M=13, N=14, O=15
const COLUNA_ID = 1;            // A - ID
const COLUNA_NOME = 2;          // B - Nome
const COLUNA_EMPRESA = 3;       // C - Empresa
const COLUNA_EMAIL = 4;         // D - Email
const COLUNA_TELEFONE = 5;      // E - Telefone
const COLUNA_CATEGORIA = 6;     // F - Categoria (NOVA POSIÇÃO)
const COLUNA_ESTADO_VENDAS = 7; // G - Status (Era F, andou +1)
const COLUNA_VALOR = 8;         // H - Valor (Era G, andou +1)
const COLUNA_DATA = 9;          // I - Data Interação (Era H, andou +1)
const COLUNA_PROX_PASSO = 10;   // J - Próximo Passo (Era I, andou +1)
const COLUNA_FONTE = 11;        // K - Fonte (Era J, andou +1)
const COLUNA_ESTADO_UF = 12;    // L - Estado UF (Era K, andou +1)
const COLUNA_ANOTACAO = 13;     // M - Anotação (Era L, andou +1)
const COLUNA_INSTAGRAM = 14;    // N - Instagram (Era M, andou +1)
const COLUNA_MAPS = 15;         // O - Maps (Era N, andou +1)

// --- GATILHOS ---
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== NOME_ABA_PROSPECCAO) return;
  const range = e.range;
  
  // Atualiza data automaticamente ao mexer no Status (Agora Coluna G/7)
  if (range.getColumn() === COLUNA_ESTADO_VENDAS && range.getRow() > 1) {
    const dataCell = sheet.getRange(range.getRow(), COLUNA_DATA);
    dataCell.setValue(new Date()); 
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Gestão de Controle 🕹️')
      .addItem('📱 Abrir Painel de Gestão', 'abrirBarraLateral')
      .addItem('💬 Abrir Comunicador (Whats/Email)', 'abrirComunicador')
      .addSeparator()
      .addItem('🧹 Organizar por Status', 'organizarPorStatus')
      .addItem('📥 Processar Fechamento', 'moverLeadsFinalizados')
      .addItem('🎨 Restaurar Design', 'aplicarEstiloVisual')
      .addToUi();
}

// --- FUNÇÕES DE INTERFACE ---
function abrirBarraLateral() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Gestão de Vendas')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function abrirComunicador() {
  const html = HtmlService.createHtmlOutputFromFile('contact')
      .setTitle('Comunicador Rápido')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- FUNÇÃO CRÍTICA (BUSCA) ---
function getListaLeads() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NOME_ABA_PROSPECCAO);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 
    
    // Lê até a coluna O (15) para garantir todos os dados
    const dados = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    
    // Índices do array são base 0 (Coluna A = 0, F = 5, G = 6...)
    return dados.map(linha => ({
      id: linha[0],          // A
      texto: `${linha[1]} | ${linha[2]}`, 
      email: linha[3],       // D
      telefone: linha[4],    // E
      categoria: linha[5],   // F (Categoria está no índice 5)
      statusAtual: linha[6], // G (Status está no índice 6)
      proxPasso: linha[9],   // J (Próx Passo está no índice 9)
      instagram: linha[13]   // N (Instagram está no índice 13)
    }));
  } catch (e) {
    console.error("Erro ao listar leads: " + e.message);
    return [];
  }
}

// --- ENVIO DE EMAIL ---
function enviarEmailLead(destinatario, assunto, mensagem) {
  if (!destinatario || !destinatario.includes("@")) throw new Error("E-mail inválido.");
  if (!assunto || !mensagem) throw new Error("Assunto/Mensagem obrigatórios.");
  try {
    MailApp.sendEmail({to: destinatario, subject: assunto, body: mensagem});
    return "E-mail enviado! 📨";
  } catch (e) { throw new Error("Erro: " + e.message); }
}

// --- ATUALIZAÇÃO ---
function atualizarStatusLead(idLead, novoStatus, dataPersonalizada, novoProximoPasso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_ABA_PROSPECCAO);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("Planilha vazia.");

  const listaIds = sheet.getRange(2, COLUNA_ID, lastRow - 1, 1).getValues().flat();
  const index = listaIds.findIndex(id => String(id) === String(idLead));
  
  if (index === -1) throw new Error("Lead não encontrado.");
  const linhaReal = index + 2;
  
  if (novoStatus) sheet.getRange(linhaReal, COLUNA_ESTADO_VENDAS).setValue(novoStatus);
  if (novoProximoPasso) sheet.getRange(linhaReal, COLUNA_PROX_PASSO).setValue(novoProximoPasso);
  
  let dataFinal = new Date(); 
  if (dataPersonalizada) dataFinal = new Date(dataPersonalizada);
  sheet.getRange(linhaReal, COLUNA_DATA).setValue(dataFinal);
  
  return `Lead atualizado!`;
}

// --- SALVAR NOVO LEAD (ORDEM ATUALIZADA) ---
function processarFormulario(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_ABA_PROSPECCAO);
  const lastRow = sheet.getLastRow();
  let novoId = 1;
  
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const numeros = ids.filter(id => !isNaN(id) && id !== "");
    if (numeros.length > 0) novoId = Math.max(...numeros) + 1;
  }

  // ARRAY DE DADOS ALINHADO COM A NOVA ORDEM (F = Categoria)
  const novaLinha = [
    novoId,                    // A - ID
    dados.nome,                // B - Nome
    dados.empresa,             // C - Empresa
    dados.email,               // D - Email
    dados.telefone,            // E - Telefone
    dados.categoria,           // F - Categoria (INSERIDA AQUI)
    "Prospecção",              // G - Status (Padrão)
    dados.valor,               // H - Valor
    new Date(),                // I - Data
    dados.proximoPasso,        // J - Próximo Passo
    dados.fonte,               // K - Fonte
    dados.cidadeEstado,        // L - Estado
    dados.oQuePrecisa,         // M - Anotação
    dados.linkInstagram,       // N - Instagram
    dados.linkGmb              // O - Maps
  ];

  sheet.appendRow(novaLinha);
  return true;
}

// --- GESTÃO ---
function organizarPorStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_ABA_PROSPECCAO);
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  // Ordena pela coluna G (7)
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({column: COLUNA_ESTADO_VENDAS, ascending: true});
  return "Tabela organizada!";
}

function moverLeadsFinalizados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaOrigem = ss.getSheetByName(NOME_ABA_PROSPECCAO);
  const abaClientes = ss.getSheetByName(NOME_ABA_CLIENTES);
  const abaPerdidos = ss.getSheetByName(NOME_ABA_PERDIDOS);
  
  if (!abaClientes || !abaPerdidos) { SpreadsheetApp.getUi().alert('ERRO: Crie as abas Clientes e Perdidos.'); return; }
  
  const lastRow = abaOrigem.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Nada para processar.'); return; }

  const dados = abaOrigem.getRange(2, 1, lastRow - 1, abaOrigem.getLastColumn()).getValues();
  let paraClientes = [], paraPerdidos = [], deletar = [];

  for (let i = 0; i < dados.length; i++) {
    // Status agora está na coluna 7 (índice 6)
    const status = String(dados[i][COLUNA_ESTADO_VENDAS - 1]); 
    if (status.includes("Ganho")) {
      paraClientes.push(dados[i]); deletar.push(i + 2);
    } else if (status === "Perdido") {
      paraPerdidos.push(dados[i]); deletar.push(i + 2);
    }
  }

  if (paraClientes.length > 0) abaClientes.getRange(abaClientes.getLastRow()+1, 1, paraClientes.length, paraClientes[0].length).setValues(paraClientes);
  if (paraPerdidos.length > 0) abaPerdidos.getRange(abaPerdidos.getLastRow()+1, 1, paraPerdidos.length, paraPerdidos[0].length).setValues(paraPerdidos);

  deletar.sort((a, b) => b - a);
  for (let i of deletar) abaOrigem.deleteRow(i);
  
  SpreadsheetApp.getUi().alert(`Processado: ${paraClientes.length} Ganhos, ${paraPerdidos.length} Perdidos.`);
}

function aplicarEstiloVisual() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const fullRange = sheet.getDataRange();
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  fullRange.setFontFamily("Roboto").setVerticalAlignment("middle").setFontSize(10);
  header.setBackground("#f1f3f4").setFontColor("#202124").setFontWeight("bold").setBorder(null, null, true, null, null, null, "#188038", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}