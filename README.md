# chatGPTonGoogleSheets: ChatGPT - Interação com API do OpenAI

Este é um script para ser utilizado no Google Sheets e que possibilita a interação com a API do OpenAI.

## Funcionalidades

### Processar com ChatGPT

Ao selecionar essa opção no menu ChatGPT, o usuário é solicitado a indicar um range de células do Google Sheets que deve ser processado pelo ChatGPT. O script processará cada célula do range indicado como input para a API do OpenAI e preencherá a célula adjacente com a resposta gerada pelo modelo de language generation.

### Atualizar chave da API

Caso o usuário ainda não tenha indicado sua chave de acesso para a API do OpenAI, ao selecionar essa opção no menu ChatGPT, ele será solicitado a fazê-lo. Caso já tenha inserido a chave previamente, essa opção permitirá a sua atualização.

### Opções de modelos de language generation

O código atualmente utiliza a versão "gpt-3.5-turbo" do modelo de language generation, mas é possível atualizar o código para utilizar o modelo "gpt-4", caso desejado.

## Utilização

### Pelo Menu

Para utilizar o script, é necessário inserir a chave de acesso para a API do OpenAI na linha correspondente do código. Feito isso, o usuário pode selecionar a opção desejada no menu ChatGPT.

### Por fórmulas

Para utilizar como fórmula basta chamar a função 'chatGPT', passando como parâmetro a solitação.

## Autores

- Autor: Marcelo Masson

## Fontes

- [Documentação da API do OpenAI](https://beta.openai.com/docs/)
- [Documentação da API do Google Sheets](https://developers.google.com/apps-script/reference/spreadsheet)
