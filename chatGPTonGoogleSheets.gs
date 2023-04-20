function onOpen() {
  const ui = SpreadsheetApp.getUi();
  console.log("Criando menu...");
  
  // Adiciona opção de menu "ChatGPT"
  ui.createMenu("ChatGPT")
    .addItem("Processar com ChatGPT", "processRangeWithChatGPT")
    .addItem("Atualizar chave da API", "updateOpenAIApiKey")
    .addItem("Atualizar modelo", "updateOpenAIModel")
    .addToUi();
  
  console.log("Menu criado com sucesso!");

  // Verifica se a chave da API está definida
  const scriptProperties = PropertiesService.getScriptProperties();
  let apiKey = scriptProperties.getProperty("openAIApiKey");

  if (!apiKey) {
    console.log("API key não encontrada.");
    const ui = SpreadsheetApp.getUi();
    
    // Pede ao usuário para inserir a chave da API
    const response = ui.prompt(
      "Insira a chave da API do OpenAI:",
      ui.ButtonSet.OK_CANCEL
    );
    
    // Atualiza a chave da API se o usuário inserir uma
    if (response.getSelectedButton() === ui.Button.OK) {
      apiKey = response.getResponseText();
      scriptProperties.setProperty("openAIApiKey", apiKey);
      console.log("API key atualizada com sucesso!");
    } else {
      ui.alert(
        "Você precisará definir a chave da API do OpenAI para usar este script."
      );
      console.log("API key não definida.");
      return;
    }
  } else {
    console.log("API key encontrada.");
  }
}

function updateOpenAIApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let apiKey = scriptProperties.getProperty("openAIApiKey");
  
  const ui = SpreadsheetApp.getUi();
  
  // Pede ao usuário para inserir a nova chave da API
  const response = ui.prompt(
    "Atualizar Key",
    "Insira a nova chave da API do OpenAI:",
    ui.ButtonSet.OK_CANCEL
  );
  
  // Atualiza a chave da API se o usuário inserir uma
  if (response.getSelectedButton() === ui.Button.OK) {
    apiKey = response.getResponseText();
    scriptProperties.setProperty("openAIApiKey", apiKey);
    ui.alert("A chave da API do OpenAI foi atualizada com sucesso!");
  }
}

function updateOpenAIModel() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const model = scriptProperties.getProperty("openAIModel");

  // Pede ao usuário para inserir o novo modelo dentre as opções
  const response = ui.prompt(
    "Configure o modelo",
    "Insira o modelo desejado entre as opções:\n\n1. gpt-4\n2. gpt-4-0314\n3. gpt-4-32k\n4. gpt-4-32k-0314\n5. gpt-3.5-turbo (Padrão)\n6. gpt-3.5-turbo-0301\n\n",
    ui.ButtonSet.OK_CANCEL
  );
  
  // Atualiza a chave da API se o usuário inserir uma
  if (response.getSelectedButton() === ui.Button.OK) {
    model = response.getResponseText();
    scriptProperties.setProperty("openAIModel", model);
    ui.alert("O modelo do chatGPT foi atualizado com sucesso!");
  }
}

function chatGPT(prompt) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty("openAIApiKey"); 
  const model = scriptProperties.getProperty("openAIModel") || "gpt-3.5-turbo"; 
  const url = 'https://api.openai.com/v1/chat/completions';
  const options = {
    "method": 'post',
    "contentType": 'application/json',
    "muteHttpExceptions": true,
    "timeout": 60000, // 60 segundos
    "headers": {
      'Authorization': 'Bearer ' + apiKey,
    },
    "payload": JSON.stringify({
      "model": model, //gpt-3.5-turbo OU gpt-4 
      "messages": [{ "role": "user", "content": prompt }]
    }),
  };
  
  // Adiciona log do prompt enviado ao ChatGPT
  console.log('Enviando prompt: ' + prompt);
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  console.log(data)
  const text = data.choices[0].message.content.trim();
  
  // Adiciona log da resposta recebida do ChatGPT
  console.log('Resposta recebida: ' + text);
  
  return text;
}

function dallE(prompt, numberImages) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty("openAIApiKey");
  const url = 'https://api.openai.com/v1/images/generations';
  const options = {
    "method": 'post',
    "contentType": 'application/json',
    "headers": {
      'Authorization': 'Bearer ' + apiKey,
    },
    "payload": JSON.stringify({
      "prompt": prompt,
      "n": numberImages || 1
    }),
  };
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  console.log(data)
  const text = data.data[0].url;
  return text;
}

function processRangeWithChatGPT() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Pede ao usuário para inserir o range dos dados a serem processados
  const rangeString = ui.prompt("Insira o range dos dados a serem processados:", ui.ButtonSet.OK_CANCEL).getResponseText();
  const range = spreadsheet.getRange(rangeString);
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  SpreadsheetApp.getActiveSpreadsheet().toast("Iniciando processamento!", null, null);
  
  let count = 0;
  for (let i = 1; i <= numRows; i++) {
    for (let j = 1; j <= numCols; j++) {
      const cell = range.getCell(i, j);
      const value = cell.getValue();
      const result = chatGPT(value);

      cell.offset(0, 1).setValue(result);
      
      // Atualiza o toast com o progresso
      count++;
      const percentComplete = Math.floor((count / (numRows * numCols)) * 100);
      SpreadsheetApp.getActiveSpreadsheet().toast("Processando... " + percentComplete + "% concluído", null, null);
    }
  }

  // Exibe a mensagem de conclusão
  SpreadsheetApp.getActiveSpreadsheet().toast("Processamento concluído!", null, null);
}
