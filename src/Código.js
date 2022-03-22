/**
 * Abre o popup com o html especificado
 * @param {string} html Nome do html desejado
 * @param {string} titulo Nome do título do modal
 * @param {number} width valor da largura
 * @param {number} height valor da altura
 */
function AbrirPopup(html, titulo, width, height) {
  var ui = HtmlService.createTemplateFromFile(html).evaluate();
  ui.setWidth(width).setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(ui, titulo);
}

/**
 * Include para o HTML
 * @param {string} filename Nome do HTML para importar na página
 * @returns html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

/**
 * Pega a spreadsheet pelo nome da aba
 * @param {string} nome Nome da aba desejada
 * @returns spreadsheet
 */
function spreadsheet(nome) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
}

// Call Popups
function AdicionarNovaDespesa() {
  AbrirPopup('despesa', 'Adicionar Nova Despesa', 400, 600);
}

function AdicionarNovaReceita() {
  AbrirPopup('receita', 'Adicionar Nova Receita', 400, 600);
}

function AdicionarTransferencia() {
  AbrirPopup('transferencia', 'Adicionar Nova Transferência', 400, 600);

}



class Money {
  constructor() {
    this.ss_Contas = spreadsheet("Contas");
    this.ss_Programacao = spreadsheet("Programacao");
    this.ss_Categorias = spreadsheet("Categorias");
    this.ss_Lancamentos = spreadsheet("Lançamentos");
    this.contas = this.ss_Programacao.getRange('AA3').getValue() + 1;
    this.cartoes = this.ss_Programacao.getRange('AA4').getValue() + 1;
    this.gastosEssenciais = this.ss_Programacao.getRange('AA5').getValue() + 1;
    this.gastos = this.ss_Programacao.getRange('AA6').getValue() + 1;
    this.receita = this.ss_Programacao.getRange('AA7').getValue() + 1;
  }
}

let money = new Money();


// -------------------
// LEITORES
// -------------------
function lerDadosNaSpreadsheet(){
  return {
    contas: money.ss_Contas.getRange("A2:A" + money.contas).getValues(),
    cartoes: money.ss_Contas.getRange("F2:F" + money.cartoes).getValues(),
    gastosEssenciais: money.ss_Categorias.getRange("A2:A" + money.gastosEssenciais).getValues().sort(),
    gastos: money.ss_Categorias.getRange("C2:C" + money.gastos).getValues().sort(),
    receitas: money.ss_Categorias.getRange("E2:E" + money.receita).getValues().sort()
  }
}

// -------------------------
// Escrever
// -------------------------

function escreverDados(data) {

  // Gera uma nova linha no topo
  money.ss_Lancamentos.insertRowBefore(2);

  // Fórmula
  money.ss_Lancamentos.getRange("F2").setFormula("=E2");

  // Copia as fórmulas para o novo campo
  money.ss_Lancamentos.getRange("H3").copyTo(money.ss_Lancamentos.getRange("H2"));

  // Adiciona os valores no spreadsheet
  money.ss_Lancamentos.getRange("A2:E2").setValues([
    [data.data, data.descricao, data.categoria, data.conta, data.valor]
  ])
}


function escreverTransferencia(data) {
  // Insere uma linha no topo
  money.ss_Lancamentos.insertRowBefore(2);

  // Copia as fórmulas
  money.ss_Lancamentos.getRange("F2").setFormula("=C2");
  money.ss_Lancamentos.getRange("H3").copyTo(money.ss_Lancamentos.getRange("H2"));

  // Insere os valores
  money.ss_Lancamentos.getRange("A2:E2").setValues([
    [data.data, `${data.contaSaida}->${data.conta}`, data.categoria, data.contaSaida, (data.valor * (-1))]
  ]);

  // Insere uma linha no topo
  money.ss_Lancamentos.insertRowBefore(2);

  // Copia as fórmulas
  money.ss_Lancamentos.getRange("F2").setFormula("=C2");
  money.ss_Lancamentos.getRange("H3").copyTo(money.ss_Lancamentos.getRange("H2"));

  // Insere os valores
  money.ss_Lancamentos.getRange("A2:E2").setValues([
    [data.data, `${data.contaSaida}->${data.conta}`, data.categoria, data.conta, data.valor]
  ]);
}

function escreverCategoria(data) {
  if (data.tipo == "Gastos_Essenciais") {
    money.ss_Categorias.getRange("A" + (money.gastosEssenciais + 1)).setValue(data.categoria);
  } else if (data.tipo == "Gastos") {
    money.ss_Categorias.getRange("C" + (money.gastos + 1)).setValue(data.categoria);
  } else if (data.tipo == "Receitas") {
    money.ss_Categorias.getRange("E" + (money.receita + 1)).setValue(data.categoria);
  }
}

function escreverConta(data) {
  if (data.tipo == "Bancos") {
    money.ss_Contas.getRange("A" + (money.contas + 1)).setValue(data.categoria);
  } else if (data.tipo == "Cartões") {
    money.ss_Contas.getRange("F" + (money.cartoes + 1)).setValue(data.categoria);
  }
}