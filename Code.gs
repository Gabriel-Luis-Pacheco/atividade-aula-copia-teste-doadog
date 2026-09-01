/**
 * ====================================================================
 * PROJETO: DOANDO CACHORROS (Inspirado no DoaDog)
 * BACKEND - GOOGLE APPS SCRIPT (Code.gs)
 * ====================================================================
 * 
 * Este arquivo processa requisições web (doGet) e lê os dados da planilha
 * do Google Sheets de forma dinâmica, adaptando-se aos nomes das colunas.
 */

// CONFIGURAÇÃO: Se o script estiver DENTRO da planilha (Script Vinculado),
// deixe SPREADSHEET_ID como "" (vazio).
// Se for um script independente (Standalone), cole o ID da sua planilha aqui.
var SPREADSHEET_ID = ""; 

// Nome aproximado ou preferencial da aba de respostas
var PREFERRED_SHEET_NAME = "Cadastro de Cachorro (respostas)";

/**
 * Função principal chamada quando o site é acessado via URL do Web App.
 */
function doGet() {
  try {
    var template = HtmlService.createTemplateFromFile('Index');
    return template.evaluate()
      .setTitle('Doando Cachorros — Adoção Responsável')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:30px;color:#c53030;">' +
      '<h2>Erro ao carregar o aplicativo</h2>' +
      '<p>' + error.message + '</p>' +
      '</div>'
    );
  }
}

/**
 * Função auxiliar para incluir arquivos HTML modulares (Styles.html, Scripts.html)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Normaliza um texto para busca de cabeçalhos (remove acentos, pontuações, espaços e caixa alta).
 */
function normalizeHeader(str) {
  if (!str) return "";
  return str.toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

/**
 * Converte links do Google Drive em links diretos para exibição em <img>
 */
function formatImageUrl(url) {
  if (!url || typeof url !== "string") return "";
  url = url.trim();

  // Se já for link direto de imagem comum
  if (url.match(/\.(jpeg|jpg|gif|png|webp|svg)($|\?)/i)) {
    return url;
  }

  // Tratamento especial para links do Google Drive
  // Formatos comuns:
  // - https://drive.google.com/file/d/FILE_ID/view?usp=sharing
  // - https://drive.google.com/open?id=FILE_ID
  // - https://drive.google.com/uc?id=FILE_ID
  var driveMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || 
                   url.match(/id=([a-zA-Z0-9_-]+)/) ||
                   url.match(/\/d\/([a-zA-Z0-9_-]+)/);

  if (driveMatch && driveMatch[1]) {
    var fileId = driveMatch[1];
    // Link direto de miniatura de alta resolução do Google Drive
    return "https://lh3.googleusercontent.com/d/" + fileId;
  }

  return url;
}

/**
 * Obtém a planilha ativa ou abre pelo ID configurado.
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  }
  
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {
    // Caso ocorra em standalone sem ID configurado
  }
  
  throw new Error("Planilha não configurada. Configure a variável SPREADSHEET_ID no Code.gs ou abra as Extensões > Apps Script diretamente de dentro da planilha.");
}

/**
 * Encontra a aba correta com os dados dos cães.
 */
function findTargetSheet(ss) {
  var sheets = ss.getSheets();
  if (sheets.length === 0) throw new Error("A planilha não possui abas.");

  // 1. Tenta encontrar por correspondência com o nome preferido
  var normPreferred = normalizeHeader(PREFERRED_SHEET_NAME);
  for (var i = 0; i < sheets.length; i++) {
    if (normalizeHeader(sheets[i].getName()).indexOf(normPreferred) !== -1 ||
        normPreferred.indexOf(normalizeHeader(sheets[i].getName())) !== -1) {
      return sheets[i];
    }
  }

  // 2. Tenta encontrar qualquer aba que contenha palavras-chave como 'resposta', 'cadastro', 'caes', 'cachorro', 'dog'
  for (var j = 0; j < sheets.length; j++) {
    var name = normalizeHeader(sheets[j].getName());
    if (name.indexOf("resposta") !== -1 || 
        name.indexOf("cadastro") !== -1 || 
        name.indexOf("cao") !== -1 || 
        name.indexOf("cachorro") !== -1 || 
        name.indexOf("dog") !== -1) {
      return sheets[j];
    }
  }

  // 3. Retorna a primeira aba caso não encontre
  return sheets[0];
}

/**
 * Identifica o papel/tipo de cada coluna através do nome do cabeçalho.
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
    ignorar: [] // Colunas administrativas como carimbo de data/hora, emails internos de login
  };

  rawHeaders.forEach(function(header, index) {
    var norm = normalizeHeader(header);
    
    // Ignorar carimbo de data/hora ou metadados de formulário desnecessários para a vitrine
    if (norm.indexOf("carimbodedata") !== -1 || norm.indexOf("timestamp") !== -1) {
      mapping.ignorar.push(index);
      return;
    }

    // Foto / Imagem
    if (mapping.foto === -1 && (norm.indexOf("foto") !== -1 || norm.indexOf("imagem") !== -1 || norm.indexOf("fotoouimagem") !== -1 || norm.indexOf("fotododog") !== -1 || norm.indexOf("linkdafoto") !== -1 || norm.indexOf("anexo") !== -1 || norm.indexOf("url") !== -1)) {
      mapping.foto = index;
    }
    // Nome do Cachorro
    else if (mapping.nome === -1 && (norm.indexOf("nomedocao") !== -1 || norm.indexOf("nomedocachorro") !== -1 || norm.indexOf("nomedoanimal") !== -1 || norm === "nome" || norm.indexOf("nome") !== -1)) {
      mapping.nome = index;
    }
    // Sexo / Gênero
    else if (mapping.sexo === -1 && (norm.indexOf("sexo") !== -1 || norm.indexOf("genero") !== -1 || norm.indexOf("machooufemea") !== -1)) {
      mapping.sexo = index;
    }
    // Porte
    else if (mapping.porte === -1 && (norm.indexOf("porte") !== -1 || norm.indexOf("tamanho") !== -1)) {
      mapping.porte = index;
    }
    // Idade
    else if (mapping.idade === -1 && (norm.indexOf("idade") !== -1 || norm.indexOf("faixaetaria") !== -1 || norm.indexOf("meses") !== -1 || norm.indexOf("anos") !== -1)) {
      mapping.idade = index;
    }
    // Raça
    else if (mapping.raca === -1 && (norm.indexOf("raca") !== -1 || norm.indexOf("tipo") !== -1)) {
      mapping.raca = index;
    }
    // Cidade / Localização / Estado
    else if (mapping.cidade === -1 && (norm.indexOf("cidade") !== -1 || norm.indexOf("localizacao") !== -1 || norm.indexOf("local") !== -1 || norm.indexOf("bairro") !== -1 || norm.indexOf("municipio") !== -1 || norm.indexOf("uf") !== -1 || norm.indexOf("estado") !== -1)) {
      mapping.cidade = index;
    }
    // Descrição / História / Sobre
    else if (mapping.descricao === -1 && (norm.indexOf("descricao") !== -1 || norm.indexOf("historia") !== -1 || norm.indexOf("sobre") !== -1 || norm.indexOf("comportamento") !== -1 || norm.indexOf("personalidade") !== -1 || norm.indexOf("resumo") !== -1)) {
      mapping.descricao = index;
    }
    // Status / Situação (Disponível, Adotado, Em processo)
    else if (mapping.status === -1 && (norm.indexOf("status") !== -1 || norm.indexOf("situacao") !== -1 || norm.indexOf("disponibilidade") !== -1 || norm.indexOf("estadodeadocao") !== -1)) {
      mapping.status = index;
    }
    // Observações / Cuidados / Saúde / Vacinas
    else if (mapping.observacoes === -1 && (norm.indexOf("observacao") !== -1 || norm.indexOf("observacoes") !== -1 || norm.indexOf("cuidados") !== -1 || norm.indexOf("saude") !== -1 || norm.indexOf("castrado") !== -1 || norm.indexOf("vacinado") !== -1 || norm.indexOf("vermifugado") !== -1)) {
      mapping.observacoes = index;
    }
    // Contato / Telefone / WhatsApp para Adoção (apenas se explícito)
    else if (mapping.contato === -1 && (norm.indexOf("contatoparadocao") !== -1 || norm.indexOf("whatsapp") !== -1 || norm.indexOf("telefoneparacontato") !== -1 || norm.indexOf("contatodivulgacao") !== -1 || norm.indexOf("responsavel") !== -1 || norm.indexOf("telefone") !== -1)) {
      mapping.contato = index;
    }
  });

  return mapping;
}

/**
 * Função principal consumida pelo front-end via google.script.run.getDogs()
 * Lê os dados da planilha e retorna uma lista estruturada de cachorros.
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

    // Pega todos os dados da planilha de uma única vez (otimização de performance)
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var rawHeaders = data[0];
    var mapping = mapHeaderColumns(rawHeaders);

    var dogs = [];

    // Itera pelas linhas de dados (a partir da linha 2)
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      
      // Verifica se a linha inteira está vazia
      var isRowEmpty = row.every(function(cell) {
        return cell === null || cell === undefined || cell.toString().trim() === "";
      });
      if (isRowEmpty) continue;

      // Extrai campos com fallback inteligente
      var nome = mapping.nome !== -1 ? String(row[mapping.nome] || "").trim() : "";
      
      // Se não encontrou coluna de nome, pega a primeira coluna de texto com conteúdo
      if (!nome) {
        for (var c = 0; c < row.length; c++) {
          if (mapping.ignorar.indexOf(c) === -1 && row[c] && typeof row[c] === 'string' && row[c].trim().length > 0) {
            nome = row[c].trim();
            break;
          }
        }
      }
      
      // Se mesmo assim não tiver identificação, usa "Amiguinho sem nome"
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

      // Monta objeto de detalhes com todos os campos existentes da planilha
      // para exibir no modal detalhado
      var outrosDetalhes = [];
      for (var colIdx = 0; colIdx < rawHeaders.length; colIdx++) {
        var headerTitle = String(rawHeaders[colIdx] || "").trim();
        var cellVal = String(row[colIdx] || "").trim();
        
        // Pula colunas ignoradas, fotos brutas ou campos vazios
        if (mapping.ignorar.indexOf(colIdx) !== -1 || colIdx === mapping.foto || !headerTitle || !cellVal) {
          continue;
        }

        // Adiciona aos detalhes extras se não for o nome básico
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
      error: err.message || "Erro desconhecido ao ler a planilha."
    };
  }
}
