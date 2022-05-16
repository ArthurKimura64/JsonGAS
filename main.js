var planilha = SpreadsheetApp.getActiveSpreadsheet();
var paginaArmazenada = planilha.getSheetByName("Support_News");
var linkArmazenado = paginaArmazenada.getRange(1, 1).getValue();

function main() {
  let array = acharLinkGED();
  link = array[0];
  ged = array[1];
  var existeLinkArmazenado = true;
  if (linkArmazenado != "") {
    if (linkArmazenado != link) {
      var gedArmazenado = JSON.parse(UrlFetchApp.fetch(linkArmazenado));
    }
    else {
      console.log("Nenhuma Mudança");
      return null;
    }
  } else { existeLinkArmazenado = false; }
  criarPaginas(ged, gedArmazenado, existeLinkArmazenado)
  if (existeLinkArmazenado) {
    if (link != linkArmazenado) {
      paginaArmazenada.getRange(1, 1).setValue(link);
    }
  }
  else {
    paginaArmazenada.getRange(1, 1).setValue(link);
  }
}
function acharLinkGED() {   //Procura o link do arquivo do GED e as informações dentro do mesmo
  let folder = DriveApp.getFolderById("13XguHM4TMe5309_dXPdB9kAzZbPUqAWN");
  let arquivo = folder.getFilesByName("index.txt").next().getBlob().getDataAsString();
  let index = JSON.parse(arquivo);
  var link = index[0].Url;
  console.log(link);
  let stringGED = UrlFetchApp.fetch(link);
  var ged = JSON.parse(stringGED);
  return Array(link, ged);
}
function criarPaginas(ged, gedArmazenado, existeLinkArmazenado) { //cria páginas na planilha
  let keys = Object.keys(ged);
  for (var contador in keys) {
    if (typeof ged[keys[contador]] === "object") {
      if (!Array.isArray(ged[keys[contador]])) {    //Se objeto não for Array, converte para Array
        ged[keys[contador]] = [ged[keys[contador]]]
      }
      if (ged[keys[contador]].length > 1) {
        //cria página caso não tenha sido criado
        if (planilha.getSheetByName(keys[contador]) == null) {
          planilha.insertSheet(keys[contador]);
        }
        if (existeLinkArmazenado) {
          if (verificarMudancas(ged[keys[contador]], gedArmazenado, keys[contador])) { continue; }
        }
        console.log("Mudança detectada em " + keys[contador])
        chaves = gerarChaves(ged[keys[contador]], keys[contador])
        inserirDados(ged[keys[contador]], gedArmazenado, keys[contador], chaves, existeLinkArmazenado)
      }
    }
  }
}
function gerarChaves(dados, nomePagina) { //Verifica chaves existentes e os insere na planilha caso tenham chaves novas;
  if (planilha.getSheetByName(nomePagina).getRange(1, 1).getValue() != "") {
    var chaves = planilha.getSheetByName(nomePagina).getRange(1, 1, 1, planilha.getSheetByName(nomePagina).getDataRange().getLastColumn()).getValues()[0];
  }
  else {
    var chaves = Object.keys(dados[dados.length - 1]);
    if (chaves.length != 0) { planilha.getSheetByName(nomePagina).getRange(1, 1, 1, chaves.length).setValues([chaves]); }
    else { var chaves = []; }
  }
  for (var numeroDeLinhas in dados) {
    var chavesLinha = Object.keys(dados[numeroDeLinhas]);
    for (var chave in chavesLinha) {
      if (chaves.indexOf(chavesLinha[chave]) == -1) {
        chaves.push(chavesLinha[chave]);
        planilha.getSheetByName(nomePagina).getRange(1, chaves.length).setValue(chavesLinha[chave]);
      }
    }
  }
  return chaves;
}
function inserirDados(dados, gedArmazenado, nomePagina, chaves, existeLinkArmazenado) { //insere os dados
  let passouAqui = true;
  MailApp.sendEmail("nendgames@gmail.com", "Houve mudança em " + nomePagina, "")
  for (var numeroDeLinhas = 0; numeroDeLinhas < dados.length; numeroDeLinhas++) {
    if (passouAqui && existeLinkArmazenado) {   //verifica em que ponto pode começar, para salvar tempo.
      if (gedArmazenado[nomePagina] != undefined) {
        if (typeof gedArmazenado[nomePagina] === "object") {
          if (!Array.isArray(gedArmazenado[nomePagina])) {
            gedArmazenado[nomePagina] = [gedArmazenado[nomePagina]];
          }
        }
      }
      if (nomePagina in gedArmazenado) {
        if (JSON.stringify(dados.slice(0, Math.ceil((dados.length - 1) / 2))) ==
          JSON.stringify(gedArmazenado[nomePagina].slice(0, Math.ceil((dados.length - 1) / 2)))) {
          if (JSON.stringify(dados.slice(0, Math.ceil((dados.length - 1) * 3 / 4))) ==
            JSON.stringify(gedArmazenado[nomePagina].slice(0, Math.ceil((dados.length - 1) * 3 / 4)))) {
            if (JSON.stringify(dados.slice(0, Math.ceil((dados.length - 1) * 7 / 8))) ==
              JSON.stringify(gedArmazenado[nomePagina].slice(0, Math.ceil((dados.length - 1) * 7 / 8)))) {
              console.log("Caiu no 7/8");
              numeroDeLinhas = Math.ceil((dados.length - 1) * 7 / 8);
            }
            else {
              console.log("Caiu no 3/4");
              numeroDeLinhas = Math.ceil((dados.length - 1) * 3 / 4);
            }
          }
          else {
            if (JSON.stringify(dados.slice(0, Math.ceil((dados.length - 1) * 5 / 8))) ==
              JSON.stringify(gedArmazenado[nomePagina].slice(0, Math.ceil((dados.length - 1) * 5 / 8)))) {
              console.log("Caiu no 5/8");
              numeroDeLinhas = Math.ceil((dados.length - 1) * 5 / 8);
            }
            else {
              console.log("Caiu no 1/2");
              numeroDeLinhas = Math.ceil((dados.length - 1) / 2);
            }
          }
        }
        else {
          if (JSON.stringify(dados.slice(0, Math.ceil((dados.length - 1) * 1 / 4))) ==
            JSON.stringify(gedArmazenado[nomePagina].slice(0, Math.ceil((dados.length - 1) * 1 / 4)))) {
            if (JSON.stringify(dados.slice(0, Math.ceil((dados.length - 1) * 3 / 8))) ==
              JSON.stringify(gedArmazenado[nomePagina].slice(0, Math.ceil((dados.length - 1) * 3 / 8)))) {
              console.log("Caiu no 3/8");
              numeroDeLinhas = Math.ceil((dados.length - 1) * 3 / 8);
            }
            else {
              console.log("Caiu no 1/4");
              numeroDeLinhas = Math.ceil((dados.length - 1) * 1 / 4);
            }
          }
          else {
            if (JSON.stringify(dados.slice(0, Math.ceil((dados.length - 1) * 1 / 8))) ==
              JSON.stringify(gedArmazenado[nomePagina].slice(0, Math.ceil((dados.length - 1) * 1 / 8)))) {
              console.log("Caiu no 1/8");
              numeroDeLinhas = Math.ceil((dados.length - 1) * 1 / 8);
            }
          }
        }
      }
      if (planilha.getSheetByName(nomePagina).getLastRow() > 2) {
        planilha.getSheetByName(nomePagina)
          .getRange(numeroDeLinhas + 2, 1, planilha.getSheetByName(nomePagina).getLastRow() - numeroDeLinhas + 1, planilha.getSheetByName(nomePagina).getLastColumn())
          .setValue("");
      }
    }
    if (!existeLinkArmazenado && passouAqui && planilha.getSheetByName(nomePagina).getLastRow() != 1) {
      numeroDeLinhas = planilha.getSheetByName(nomePagina).getLastRow() - 2;
      if (dados.length <= numeroDeLinhas) { numeroDeLinhas = dados.length - 1 }
    }
    passouAqui = false
    if (dados[numeroDeLinhas] != "undefined") {
      let chavesLinha = Object.keys(dados[numeroDeLinhas]);
      for (var chave in chavesLinha) {
        let coluna = chaves.indexOf(chavesLinha[chave]) + 1;
        if (typeof dados[numeroDeLinhas][chavesLinha[chave]] === "object") {
          dados[numeroDeLinhas][chavesLinha[chave]] = JSON.stringify(dados[numeroDeLinhas][chavesLinha[chave]]);
        }
        //Formata corretamente o formato das datas
        if (String(dados[numeroDeLinhas][chavesLinha[chave]]).match(/\d+\-\d+\-\d+/) !== null ||
          String(dados[numeroDeLinhas][chavesLinha[chave]]).match(/\d+\/\d+\/\d+/) !== null) {
          if (dados[numeroDeLinhas][chavesLinha[chave]].match(/\d+\-\d+\-\d+/) != null) { var formatoData = /\d+\-\d+\-\d+/; }
          else { var formatoData = /\d+\/\d+\/\d+/; }
          X = dados[numeroDeLinhas][chavesLinha[chave]].match(formatoData)[0];
          planilha.getSheetByName(nomePagina)
            .getRange(numeroDeLinhas + 2, coluna)
            .setValue(new Date(X).toLocaleDateString("pt-BR"));
        }
        else {
          planilha.getSheetByName(nomePagina)
            .getRange(numeroDeLinhas + 2, coluna)
            .setValue(String(dados[numeroDeLinhas][chavesLinha[chave]]));
        }
      }
    }
  }
}
function verificarMudancas(dados, gedArmazenado, nomePagina) {
  if (gedArmazenado[nomePagina] !== undefined) {
    if (typeof gedArmazenado[nomePagina] === "object") {
      if (!Array.isArray(gedArmazenado[nomePagina])) {
        gedArmazenado[nomePagina] = [gedArmazenado[nomePagina]];
      }
    }
    if (JSON.stringify(gedArmazenado[nomePagina]) == JSON.stringify(dados)) { return true; }
    else { return false; }
  }
  else { return false; }
}


function localization() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  if (planilha.getSheetByName("Localization") == null) {
    planilha.insertSheet("Localization");
  }
  var pagina = planilha.getSheetByName("Localization");
  folder = DriveApp.getFolderById("1jAzDv-ozNr9YdnuM42eOaZ_-soGN3mXo")
    .getFilesByName("index.txt")
    .next()
    .getBlob()
    .getDataAsString();
  index = JSON.parse(folder);
  Link = index[0].Url;
  Logger.log(Link);
  var StringLocalization = UrlFetchApp.fetch(Link);
  var Local = JSON.parse(StringLocalization);
  var dados = Local["terms"];
  var linguagem = index[0]["Id"];
  var keys = [];
  if (pagina.getRange(1, 2).getValue() == "") {
    keys.push("data");
    pagina.getRange(1, keys.length).setValue(keys[keys.length - 1]);
    keys.push(linguagem);
    pagina.getRange(1, keys.length).setValue(keys[keys.length - 1]);
  } else {
    keys = pagina.getRange(1, 1, 1, pagina.getLastColumn()).getValues()[0];
  }
  var coluna = keys.indexOf(linguagem) + 1;
  var informacoes = pagina
    .getRange(1, coluna, pagina.getLastRow(), 1)
    .getValues();
  for (valores in informacoes) {
    if (informacoes[valores] == "") {
      informacoes[valores] = null;
    }
  }
  informacoes = informacoes.filter(function (x) {
    return x;
  });
  if (pagina.getLastRow() * 2 == dados.length) {
  } else {
    for (var i = (informacoes.length) * 2 - 1; i <= dados.length - 1; i++) {
      if (i % 2 == 0) {
        pagina.getRange(Math.floor(i / 2) + 2, 1).setValue(dados[i]);
      }
      if (i % 2 == 1) {
        if (dados[i] == "" || dados[i] == null) {
          dados[i] = "-";
        }
        pagina.getRange(Math.floor(i / 2) + 2, coluna).setValue(dados[i]);
      }
    }
  }
}
