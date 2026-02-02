/**
 * ------------------------------------------------------------------
 * CONFIGURAÇÕES GERAIS E CONSTANTES
 * ------------------------------------------------------------------
 */
const SPREADSHEET_ID = '146nlNN-ym44OjMtr0bceupaVZ-7m42rXByt8UZf0TUk';
const FOLDER_ID_IMAGENS = '1X_cAyMR85E0gnOwPJi_ychjsBjVzlNRc';
const FOLDER_ID_PDFS = '1fnlEAWr6Q2my1-P6tAU7AglEmVy-Lz6-';

/**
 * Função padrão para servir o HTML quando o usuário acessa o Web App.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Cadastro Médico - Ecossistema de Precisão')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ------------------------------------------------------------------
 * DATA FETCHING (OTIMIZAÇÃO)
 * ------------------------------------------------------------------
 */
function getDropdownData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const getColData = (sheetName, colIndex) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, colIndex, lastRow - 1, 1).getValues().flat().filter(String);
  };

  const faculdades = getColData('FACULDADES', 3).sort();
  const especialidades = getColData('ESPECIALIDADES', 1).sort();

  const sheetRes = ss.getSheetByName('RESIDÊNCIAS');
  const rawRes = sheetRes ? sheetRes.getRange(2, 1, sheetRes.getLastRow() - 1, 2).getValues() : [];
  
  const residenciasMap = {};
  rawRes.forEach(([programa, hospital]) => {
    if (programa && hospital) {
      if (!residenciasMap[programa]) residenciasMap[programa] = [];
      residenciasMap[programa].push(hospital);
    }
  });

  return {
    faculdades,
    especialidades,
    residenciasMap
  };
}

/**
 * ------------------------------------------------------------------
 * LÓGICA DE PERSISTÊNCIA E REGRAS DE NEGÓCIO
 * ------------------------------------------------------------------
 */
function processForm(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('CADASTRO');
    
    // Helper para extrair extensão do arquivo
    const getFileExtension = (filename, mimeType) => {
      if (filename && filename.includes('.')) {
        return filename.split('.').pop();
      }
      // Fallback básico se o nome não tiver extensão
      if (mimeType === 'application/pdf') return 'pdf';
      if (mimeType === 'image/jpeg') return 'jpg';
      if (mimeType === 'image/png') return 'png';
      return 'bin';
    };

    // 1. Processamento de Arquivos com Renomeação Personalizada
    let imageUrl = '';
    let pdfUrl = '';

    // Lógica para FOTO: NOME COMPLETO + NIP
    if (payload.fotoBase64) {
      const ext = getFileExtension(payload.fotoNome, payload.fotoMime);
      const novoNomeFoto = `${payload.nomeCompleto} - ${payload.nip}.${ext}`;
      imageUrl = uploadFileToDrive(payload.fotoBase64, novoNomeFoto, payload.fotoMime, FOLDER_ID_IMAGENS);
    }

    // Lógica para PDF: "CRM" + NOME COMPLETO + NIP
    if (payload.docBase64) {
      const ext = getFileExtension(payload.docNome, payload.docMime);
      const novoNomePdf = `CRM ${payload.nomeCompleto} - ${payload.nip}.${ext}`;
      pdfUrl = uploadFileToDrive(payload.docBase64, novoNomePdf, payload.docMime, FOLDER_ID_PDFS);
    }

    // 2. Cálculos
    const anoAtual = new Date().getFullYear();
    const anoNasc = payload.nascimento ? new Date(payload.nascimento).getFullYear() : 0;
    const idade = (anoNasc > 0) ? (anoAtual - anoNasc) : '';
    const naturalidade = `${payload.naturalidadeCidade || ''} - ${payload.naturalidadeUF || ''}`;

    // 3. Montagem da Linha
    const rowData = [
      payload.nomeCompleto,
      payload.nascimento,
      idade,
      naturalidade,
      payload.cpf,
      payload.celular,
      payload.email,
      payload.nip,
      payload.crm,
      payload.enderecoLogradouro,
      payload.numero,
      payload.complemento,
      payload.faculdade,
      payload.formatura,
      payload.especialidade,
      payload.vagaTrancada, // Agora passa o Boolean direto (True/False) para o Checkbox
      payload.programaResidencia || '',
      payload.hospital || '',
      imageUrl,
      pdfUrl
    ];

    sheet.appendRow(rowData);

    return { success: true, message: 'Cadastro realizado com sucesso!' };

  } catch (error) {
    return { success: false, message: 'Erro ao salvar: ' + error.message };
  }
}

/**
 * Helper para upload de arquivo Base64
 */
function uploadFileToDrive(base64Data, filename, mimeType, folderId) {
  try {
    const data = base64Data.split(',')[1] || base64Data;
    const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, filename);
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    console.error('Erro no upload: ' + e.toString());
    return 'Erro no Upload';
  }
}
