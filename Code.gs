/**
 * ====================================================================
 * PROJETO: DOANDO CACHORROS (Inspirado no DoaDog)
 * BACKEND - GOOGLE APPS SCRIPT (Code.gs)
 * ====================================================================
 */

// Se abrir o Apps Script por dentro da planilha (Extensões > Apps Script),
// pode deixar vazio (""). Se for um script independente, cole o ID aqui.
var SPREADSHEET_ID = ""; 

// Nome aproximado ou preferencial da aba de respostas
var PREFERRED_SHEET_NAME = "Cadastro de Cachorro (respostas)";

/**
 * Ponto de entrada Web App (GET)
 */
function doGet() {
  try {
    var htmlOutput = HtmlService.createTemplateFromFile('Index').evaluate();
    return htmlOutput
      .setTitle('Doando Cachorros — Uma versão demonstrativa inspirada no projeto DoaDog')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    // Fallback simples caso não use template
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Doando Cachorros — Adoção Responsável')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
}

/**
 * Função de inclusão modular caso o usuário queira separar arquivos
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return "";
  }
}

/**
 * Normaliza textos de cabeçalho para comparação flexível
 */
function normalizeText(str) {
  if (!str) return "";
  return str.toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

/**
 * Transforma links do Google Drive ou da web em links diretos para <img>
 */
function formatImageUrl(url) {
  if (!url || typeof url !== "string") return "";
  url = url.trim();

  // Se já for link direto com extensão de imagem
  if (url.match(/\.(jpeg|jpg|gif|png|webp|svg|bmp)(\?.*)?$/i)) {
    return url;
  }

  // Links do Google Drive (visualização, compartilhamento, anexo de formulário)
  var driveMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || 
                   url.match(/id=([a-zA-Z0-9_-]+)/) ||
                   url.match(/\/d\/([a-zA-Z0-9_-]+)/);

  if (driveMatch && driveMatch[1]) {
    var fileId = driveMatch[1];
    return "https://lh3.googleusercontent.com/d/" + fileId;
  }

  return url;
}

/**
 * Obtém a planilha ativa ou pelo ID configurado
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  }
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  throw new Error("Planilha não configurada. Abra o Apps Script pelo menu 'Extensões > Apps Script' da planilha ou configure o SPREADSHEET_ID no Code.gs.");
}

/**
 * Localiza a aba correta com os dados cadastrais dos cachorros
 */
function findTargetSheet(ss) {
  var sheets = ss.getSheets();
  if (sheets.length === 0) throw new Error("A planilha não possui abas.");

  var normPreferred = normalizeText(PREFERRED_SHEET_NAME);
  for (var i = 0; i < sheets.length; i++) {
    var currentNorm = normalizeText(sheets[i].getName());
    if (currentNorm.indexOf(normPreferred) !== -1 || normPreferred.indexOf(currentNorm) !== -1) {
      return sheets[i];
    }
  }

  for (var j = 0; j < sheets.length; j++) {
    var name = normalizeText(sheets[j].getName());
    if (name.indexOf("resposta") !== -1 || 
        name.indexOf("cadastro") !== -1 || 
        name.indexOf("cao") !== -1 || 
        name.indexOf("cachorro") !== -1 || 
        name.indexOf("dog") !== -1 ||
        name.indexOf("form") !== -1) {
      return sheets[j];
    }
  }

  return sheets[0];
}

/**
 * Mapeia dinamicamente os nomes das colunas aos campos do sistema
 */
function mapHeaderColumns(rawHeaders) {
  var mapping = {
    nome: -1,
    foto: -1,
    idade: -1,
    sexo: -1,
    porte: -1,
    raca: -1,
    cidade: -1,
    descricao: -1,
    status: -1,
    observacoes: -1,
    contato: -1,
    ignorar: []
  };

  rawHeaders.forEach(function(header, index) {
    var norm = normalizeText(header);
    
    // Ignorar carimbos de data/hora técnicos
    if (norm.indexOf("carimbodedata") !== -1 || norm.indexOf("timestamp") !== -1) {
      mapping.ignorar.push(index);
      return;
    }

    // Foto / Imagem / Anexo
    if (mapping.foto === -1 && (norm.indexOf("foto") !== -1 || norm.indexOf("imagem") !== -1 || norm.indexOf("anexo") !== -1 || norm.indexOf("linkdafoto") !== -1 || norm.indexOf("url") !== -1)) {
      mapping.foto = index;
    }
    // Nome do Cachorro
    else if (mapping.nome === -1 && (norm.indexOf("nomedocao") !== -1 || norm.indexOf("nomedocachorro") !== -1 || norm.indexOf("nomedoanimal") !== -1 || norm === "nome" || norm.indexOf("nome") !== -1)) {
      mapping.nome = index;
    }
    // Sexo / Gênero
    else if (mapping.sexo === -1 && (norm.indexOf("sexo") !== -1 || norm.indexOf("genero") !== -1 || norm.indexOf("macho") !== -1 || norm.indexOf("femea") !== -1)) {
      mapping.sexo = index;
    }
    // Porte / Tamanho
    else if (mapping.porte === -1 && (norm.indexOf("porte") !== -1 || norm.indexOf("tamanho") !== -1)) {
      mapping.porte = index;
    }
    // Idade / Faixa Etária
    else if (mapping.idade === -1 && (norm.indexOf("idade") !== -1 || norm.indexOf("faixaetaria") !== -1 || norm.indexOf("anos") !== -1 || norm.indexOf("meses") !== -1)) {
      mapping.idade = index;
    }
    // Raça
    else if (mapping.raca === -1 && (norm.indexOf("raca") !== -1 || norm.indexOf("tipo") !== -1 || norm.indexOf("especie") !== -1)) {
      mapping.raca = index;
    }
    // Cidade / Localização / Bairro / UF
    else if (mapping.cidade === -1 && (norm.indexOf("cidade") !== -1 || norm.indexOf("localizacao") !== -1 || norm.indexOf("local") !== -1 || norm.indexOf("bairro") !== -1 || norm.indexOf("municipio") !== -1 || norm.indexOf("uf") !== -1 || norm.indexOf("estado") !== -1)) {
      mapping.cidade = index;
    }
    // Descrição / História / Sobre / Comportamento
    else if (mapping.descricao === -1 && (norm.indexOf("descricao") !== -1 || norm.indexOf("historia") !== -1 || norm.indexOf("sobre") !== -1 || norm.indexOf("comportamento") !== -1 || norm.indexOf("personalidade") !== -1 || norm.indexOf("resumo") !== -1)) {
      mapping.descricao = index;
    }
    // Status / Situação (Disponível, Em processo, Adotado)
    else if (mapping.status === -1 && (norm.indexOf("status") !== -1 || norm.indexOf("situacao") !== -1 || norm.indexOf("disponibilidade") !== -1 || norm.indexOf("adocao") !== -1)) {
      mapping.status = index;
    }
    // Observações / Saúde / Vacinas / Castração
    else if (mapping.observacoes === -1 && (norm.indexOf("observacao") !== -1 || norm.indexOf("observacoes") !== -1 || norm.indexOf("cuidados") !== -1 || norm.indexOf("saude") !== -1 || norm.indexOf("castrado") !== -1 || norm.indexOf("vacinado") !== -1)) {
      mapping.observacoes = index;
    }
    // Contato / Telefone / WhatsApp para Adoção
    else if (mapping.contato === -1 && (norm.indexOf("contatoparadocao") !== -1 || norm.indexOf("whatsapp") !== -1 || norm.indexOf("telefone") !== -1 || norm.indexOf("contato") !== -1 || norm.indexOf("responsavel") !== -1)) {
      mapping.contato = index;
    }
  });

  return mapping;
}

/**
 * Função chamada pelo front-end através de google.script.run.getDogs()
 */
function getDogs() {
  try {
    var ss = getSpreadsheet();
    var sheet = findTargetSheet(ss);
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < 2 || lastCol < 1) {
      return {
        success: true,
        sheetName: sheet.getName(),
        total: 0,
        dogs: []
      };
    }

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var rawHeaders = data[0];
    var mapping = mapHeaderColumns(rawHeaders);

    var dogs = [];

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      
      var isRowEmpty = row.every(function(cell) {
        return cell === null || cell === undefined || cell.toString().trim() === "";
      });
      if (isRowEmpty) continue;

      var nome = mapping.nome !== -1 ? String(row[mapping.nome] || "").trim() : "";
      if (!nome) {
        for (var c = 0; c < row.length; c++) {
          if (mapping.ignorar.indexOf(c) === -1 && row[c] && typeof row[c] === 'string' && row[c].trim().length > 0) {
            nome = row[c].trim();
            break;
          }
        }
      }
      if (!nome) nome = "Amiguinho sem nome";

      var rawFoto = mapping.foto !== -1 ? String(row[mapping.foto] || "").trim() : "";
      var fotoUrl = formatImageUrl(rawFoto);

      var idade = mapping.idade !== -1 ? String(row[mapping.idade] || "").trim() : "";
      var sexo = mapping.sexo !== -1 ? String(row[mapping.sexo] || "").trim() : "";
      var porte = mapping.porte !== -1 ? String(row[mapping.porte] || "").trim() : "";
      var raca = mapping.raca !== -1 ? String(row[mapping.raca] || "").trim() : "";
      var cidade = mapping.cidade !== -1 ? String(row[mapping.cidade] || "").trim() : "";
      var descricao = mapping.descricao !== -1 ? String(row[mapping.descricao] || "").trim() : "";
      var status = mapping.status !== -1 ? String(row[mapping.status] || "").trim() : "Disponível";
      var observacoes = mapping.observacoes !== -1 ? String(row[mapping.observacoes] || "").trim() : "";
      var contato = mapping.contato !== -1 ? String(row[mapping.contato] || "").trim() : "";

      var outrosDetalhes = [];
      for (var colIdx = 0; colIdx < rawHeaders.length; colIdx++) {
        var headerTitle = String(rawHeaders[colIdx] || "").trim();
        var cellVal = String(row[colIdx] || "").trim();
        
        if (mapping.ignorar.indexOf(colIdx) !== -1 || colIdx === mapping.foto || !headerTitle || !cellVal) {
          continue;
        }

        if (colIdx !== mapping.nome) {
          outrosDetalhes.push({
            titulo: headerTitle,
            valor: cellVal
          });
        }
      }

      dogs.push({
        id: "dog_" + r,
        linha: r + 1,
        nome: nome,
        foto: fotoUrl,
        idade: idade || "Não informada",
        sexo: sexo || "Não informado",
        porte: porte || "Não informado",
        raca: raca || "SRD (Sem raça definida)",
        cidade: cidade || "Não informada",
        descricao: descricao || "Sem descrição informada.",
        status: status || "Disponível",
        observacoes: observacoes,
        contato: contato,
        detalhesCompletos: outrosDetalhes
      });
    }

    return {
      success: true,
      sheetName: sheet.getName(),
      total: dogs.length,
      dogs: dogs
    };

  } catch (err) {
    return {
      success: false,
      error: err.message || "Erro ao consultar a planilha."
    };
  }
}
