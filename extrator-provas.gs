/**
 * Script para Google Sheets que se conecta à API do Google Gemini
 * para extrair informações (nome e média) das capas de provas
 * e inserir esses dados na planilha.
 * 
 * IMPORTANTE: Insira sua chave API do Gemini diretamente na constante abaixo.
 */

// Configurações da API Gemini - INSIRA SUA CHAVE AQUI
const GEMINI_API_KEY = 'SUA_CHAVE_API_AQUI'; // ⚠️ SUBSTITUA COM SUA CHAVE REAL DO GEMINI
const GEMINI_MODEL = 'gemini-2.0-flash'; // Modelo Gemini 2.0 Flash

/**
 * Cria o menu na interface do Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Extrator de Provas')
    .addItem('Processar imagens de provas', 'processarImagens')
    .addItem('Processar imagem específica', 'processarImagemEspecifica')
    .addItem('Listar arquivos na pasta', 'listarArquivosNaPasta')
    .addToUi();
}

/**
 * Função para listar arquivos na pasta - ajuda na depuração
 */
function listarArquivosNaPasta() {
  const ui = SpreadsheetApp.getUi();
  
  // Solicitar pasta com as imagens das provas
  const response = ui.prompt(
    'Listar Arquivos',
    'Insira o ID da pasta do Google Drive:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  const folderId = response.getResponseText();
  
  try {
    // Obter a pasta
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    
    let fileList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        nome: file.getName(),
        tipo: file.getMimeType(),
        id: file.getId(),
        url: file.getUrl(),
        tamanho: Math.round(file.getSize() / 1024) + ' KB'
      });
    }
    
    // Criar uma nova planilha para mostrar os arquivos
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Limpar dados anteriores
    sheet.clear();
    
    // Criar cabeçalhos
    sheet.getRange('A1:E1').setValues([['Nome', 'Tipo', 'ID', 'URL', 'Tamanho']]);
    sheet.getRange('A1:E1').setFontWeight('bold');
    
    // Preencher com a lista de arquivos
    if (fileList.length > 0) {
      for (let i = 0; i < fileList.length; i++) {
        sheet.getRange(i + 2, 1, 1, 5).setValues([[
          fileList[i].nome,
          fileList[i].tipo,
          fileList[i].id,
          fileList[i].url,
          fileList[i].tamanho
        ]]);
      }
      
      // Ajustar largura das colunas
      sheet.autoResizeColumns(1, 5);
      
      ui.alert(`Encontrados ${fileList.length} arquivos na pasta.`);
    } else {
      ui.alert('Nenhum arquivo encontrado na pasta.');
    }
    
  } catch (e) {
    ui.alert('Erro: ' + e.toString());
    Logger.log('Erro completo: ' + e.stack);
  }
}

/**
 * Função para processar uma imagem específica
 */
function processarImagemEspecifica() {
  const ui = SpreadsheetApp.getUi();
  
  // Verificar se a chave API foi definida
  if (!GEMINI_API_KEY || GEMINI_API_KEY === 'SUA_CHAVE_API_AQUI') {
    ui.alert('ERRO: Você precisa inserir sua chave API do Gemini no código do script.');
    return;
  }
  
  // Solicitar ID do arquivo específico
  const response = ui.prompt(
    'Processar Imagem',
    'Insira o ID do arquivo no Google Drive:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  const fileId = response.getResponseText();
  
  try {
    // Obter o arquivo
    const file = DriveApp.getFileById(fileId);
    
    // Preparar a planilha
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Criar cabeçalhos se não existirem
    if (sheet.getRange('A1').getValue() === '') {
      sheet.getRange('A1:C1').setValues([['Nome do Aluno', 'Média', 'URL da Imagem']]);
      sheet.getRange('A1:C1').setFontWeight('bold');
    }
    
    // Mostrar informações do arquivo
    ui.alert(`Processando arquivo: ${file.getName()} (${file.getMimeType()}, ${Math.round(file.getSize()/1024)} KB)`);
    
    // Processar o arquivo
    try {
      const resultado = processarImagemComGemini(file);
      
      if (resultado) {
        const lastRow = sheet.getLastRow() + 1;
        sheet.getRange(lastRow, 1, 1, 3).setValues([[
          resultado.nomeAluno,
          resultado.media,
          file.getUrl()
        ]]);
        ui.alert(`Sucesso! Nome: ${resultado.nomeAluno}, Média: ${resultado.media}`);
      } else {
        ui.alert('Não foi possível extrair os dados da imagem.');
      }
    } catch (processError) {
      ui.alert(`Erro ao processar imagem: ${processError.message}`);
      Logger.log(`Erro detalhado: ${processError.stack}`);
    }
    
  } catch (e) {
    ui.alert('Erro: ' + e.toString());
    Logger.log('Erro completo: ' + e.stack);
  }
}

/**
 * Função principal que processa as imagens das provas armazenadas em uma pasta do Drive
 */
function processarImagens() {
  const ui = SpreadsheetApp.getUi();
  
  // Verificar se a chave API foi definida
  if (!GEMINI_API_KEY || GEMINI_API_KEY === 'SUA_CHAVE_API_AQUI') {
    ui.alert('ERRO: Você precisa inserir sua chave API do Gemini no código do script.');
    return;
  }
  
  // Fazer um teste rápido da API antes de continuar
  try {
    const testResult = testSimpleApiConnection();
    if (!testResult.success) {
      ui.alert(`Erro ao testar API do Gemini: ${testResult.error}`);
      return;
    }
  } catch (e) {
    ui.alert(`Erro ao conectar com a API do Gemini: ${e.toString()}`);
    return;
  }
  
  // Solicitar pasta com as imagens das provas
  const response = ui.prompt(
    'Processar Provas',
    'Insira o ID da pasta do Google Drive contendo as imagens das provas:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  const folderId = response.getResponseText();
  
  try {
    // Obter a pasta do Drive
    const folder = DriveApp.getFolderById(folderId);
    
    // Obter todos os arquivos
    const files = folder.getFiles();
    let allFiles = [];
    let contadores = {
      total: 0,
      imagens: 0,
      pdfs: 0,
      outros: 0
    };
    
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      
      allFiles.push(file);
      contadores.total++;
      
      if (mimeType === MimeType.JPEG || mimeType === MimeType.PNG) {
        contadores.imagens++;
      } else if (mimeType === MimeType.PDF) {
        contadores.pdfs++;
      } else {
        contadores.outros++;
      }
    }
    
    // Log para depuração
    Logger.log(`Encontrados ${contadores.total} arquivos: ${contadores.imagens} imagens, ${contadores.pdfs} PDFs, ${contadores.outros} outros`);
    
    // Preparar a planilha
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Criar cabeçalhos se não existirem
    if (sheet.getRange('A1').getValue() === '') {
      sheet.getRange('A1:D1').setValues([['Nome do Aluno', 'Média', 'URL da Imagem', 'Status']]);
      sheet.getRange('A1:D1').setFontWeight('bold');
    }
    
    // Processar cada arquivo
    let lastRow = sheet.getLastRow();
    let contador = 0;
    let sucessos = 0;
    
    ui.alert(`Iniciando processamento de ${allFiles.length} arquivos: ${contadores.imagens} imagens, ${contadores.pdfs} PDFs, ${contadores.outros} outros`);
    
    // Processar os arquivos da lista
    for (let i = 0; i < allFiles.length; i++) {
      const file = allFiles[i];
      const fileName = file.getName();
      const imageUrl = file.getUrl();
      const mimeType = file.getMimeType();
      
      // Log para debugging
      Logger.log(`Processando arquivo ${i+1}/${allFiles.length}: ${fileName} (${mimeType})`);
      
      // Pular arquivos que não sejam imagens 
      if (mimeType !== MimeType.JPEG && mimeType !== MimeType.PNG) {
        Logger.log(`Pulando ${fileName}: formato não suportado (${mimeType})`);
        continue;
      }
      
      // Obter dados da imagem usando a API do Gemini
      try {
        const resultado = processarImagemComGemini(file);
        
        lastRow++;
        if (resultado) {
          sheet.getRange(lastRow, 1, 1, 4).setValues([[
            resultado.nomeAluno,
            resultado.media,
            imageUrl,
            "Sucesso"
          ]]);
          sucessos++;
          Logger.log(`Sucesso: Nome=${resultado.nomeAluno}, Média=${resultado.media}`);
        } else {
          sheet.getRange(lastRow, 1, 1, 4).setValues([[
            "N/A",
            "N/A",
            imageUrl,
            "Falha na extração"
          ]]);
          Logger.log(`Falha na extração: ${fileName}`);
        }
        contador++;
      } catch (processError) {
        lastRow++;
        sheet.getRange(lastRow, 1, 1, 4).setValues([[
          "N/A",
          "N/A",
          imageUrl,
          "Erro: " + processError.message.substring(0, 100)
        ]]);
        Logger.log(`Erro ao processar ${fileName}: ${processError.message}`);
        contador++;
      }
    }
    
    ui.alert(`Processamento concluído! Foram processados ${contador} arquivos com ${sucessos} extrações bem-sucedidas.`);
    
  } catch (e) {
    ui.alert('Erro: ' + e.toString());
    Logger.log('Erro completo: ' + e.stack);
  }
}

/**
 * Processa uma imagem usando a API do Gemini
 * @param {File} file - Objeto do arquivo do Google Drive
 * @return {Object} Objeto contendo nome do aluno e média extraídos
 */
function processarImagemComGemini(file) {
  // Verificar se a chave API está configurada
  if (!GEMINI_API_KEY || GEMINI_API_KEY === 'SUA_CHAVE_API_AQUI') {
    throw new Error('Chave de API do Gemini não configurada. Insira sua chave no código.');
  }
  
  // Verificar tipo de arquivo
  const mimeType = file.getMimeType();
  const fileName = file.getName();
  
  Logger.log(`Processando ${fileName} (${mimeType})`);
  
  // Obter o blob da imagem
  const blob = file.getBlob();
  const base64Image = Utilities.base64Encode(blob.getBytes());
  
  // Log para acompanhar o progresso
  Logger.log(`Enviando ${fileName} para o Gemini (tamanho: ${Math.round(base64Image.length/1024)}KB)`);
  
  // Verificar tamanho da imagem - Gemini tem limite
  if (base64Image.length > 4 * 1024 * 1024) { // Reduzindo para 4MB para maior segurança
    Logger.log(`Imagem muito grande: ${fileName} (${Math.round(base64Image.length/1024/1024)}MB)`);
    throw new Error(`Imagem muito grande: ${Math.round(base64Image.length/1024/1024)}MB - limite de 4MB`);
  }
  
  // Criar o payload para a API do Gemini
  const payload = {
    contents: [
      {
        parts: [
          {
            text: "Você é um assistente especializado em extrair informações de documentos acadêmicos. Esta é uma folha de resposta acadêmica. Por favor, extraia o nome completo do aluno e a nota média. Retorne apenas um objeto JSON no formato {\"nomeAluno\": \"Nome completo do aluno\", \"media\": número}."
          },
          {
            inlineData: {
              mimeType: mimeType,
              data: base64Image
            }
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 150,
      topP: 0.1,
      topK: 16
    }
  };
  
  // Configurar a requisição HTTP para Gemini 2.0
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    // Aumentar o timeout para dar tempo suficiente para processamento de imagens
    timeout: 60000
  };
  
  try {
    // Fazer a requisição para a API do Gemini
    Logger.log("Enviando requisição para API do Gemini...");
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`Resposta recebida: Código ${responseCode}`);
    
    if (responseCode !== 200) {
      Logger.log(`Erro na API do Gemini: Status ${responseCode}, Resposta: ${responseText}`);
      throw new Error(`Erro na API: ${responseCode} - ${responseText.substring(0, 100)}...`);
    }
    
    // Log da resposta completa para depuração
    Logger.log(`Resposta completa: ${responseText}`);
    
    const responseData = JSON.parse(responseText);
    
    // Extrair o conteúdo da resposta
    if (responseData.candidates && responseData.candidates.length > 0) {
      const content = responseData.candidates[0].content;
      
      if (content && content.parts && content.parts.length > 0) {
        const text = content.parts[0].text;
        Logger.log(`Texto extraído: ${text}`);
        
        // Tentar extrair o JSON da resposta
        try {
          // Remover qualquer texto antes ou depois do JSON
          const jsonMatch = text.match(/\{.*\}/s);
          if (jsonMatch) {
            const jsonStr = jsonMatch[0];
            Logger.log(`JSON encontrado: ${jsonStr}`);
            
            const resultado = JSON.parse(jsonStr);
            
            // Verificar se os dados necessários estão presentes
            if (resultado.nomeAluno && resultado.media !== undefined) {
              return {
                nomeAluno: resultado.nomeAluno,
                media: parseFloat(resultado.media) // Converter para número
              };
            } else {
              Logger.log("JSON não contém os campos esperados");
              return null;
            }
          } else {
            Logger.log("Nenhum JSON encontrado na resposta");
            return null;
          }
        } catch (jsonError) {
          Logger.log('Erro ao analisar resposta JSON: ' + jsonError);
          throw new Error('Erro ao analisar resposta JSON: ' + jsonError.message);
        }
      } else {
        Logger.log("Resposta sem conteúdo de texto");
        return null;
      }
    } else {
      Logger.log("Resposta sem candidatos");
      return null;
    }
    
  } catch (e) {
    Logger.log('Erro na requisição à API: ' + e);
    Logger.log('Stack trace: ' + e.stack);
    throw e;
  }
}

/**
 * Função simples para testar apenas a conexão com a API
 */
function testSimpleApiConnection() {
  try {
    // Criar um payload de teste simples
    const testPayload = {
      contents: [
        {
          parts: [
            {
              text: "Teste de conexão. Responda apenas com 'OK'."
            }
          ]
        }
      ]
    };
    
    // URL da API para Gemini 2.0
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
    
    // Fazer uma requisição simples para testar a conexão
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(testPayload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      return { success: true };
    } else {
      const responseText = response.getContentText();
      return {
        success: false,
        error: `Código de erro ${responseCode}: ${responseText}`
      };
    }
  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}
