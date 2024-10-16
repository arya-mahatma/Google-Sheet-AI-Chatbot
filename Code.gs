function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Create the main menu
  var menu = ui.createMenu('AI Chatbot')
    .addItem('Start AI Chatbot', 'showChatSidebar')
    .addSeparator();

  // List of AI providers
  var providers = [
    'OpenAI', 'Groq', 'Together', 'Google', 'Anthropic', 'Hyperbolic', 'Mistral', 'Cerebras', 'SambaNova'
  ];

  // Create submenus for each provider
  providers.forEach(function(provider) {
    var providerMenu = ui.createMenu(provider)
      .addItem('Set API Key', 'set' + provider.replace(/\s+/g, '') + 'ApiKey')
      .addItem('View API Key', 'view' + provider.replace(/\s+/g, '') + 'ApiKey')
      .addItem('Clear API Key', 'clear' + provider.replace(/\s+/g, '') + 'ApiKey');
    
    menu.addSubMenu(providerMenu);
  });

  // Add the menu to the UI
  menu.addToUi();
}

function showChatSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('AI Chatbot')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function formatKeyName(provider) {
  return `${provider.toUpperCase().replace(/\s+/g, '_')}_API_KEY`;
}

function chatbot(provider, message, conversationHistory = []) {
  const keyName = formatKeyName(provider);
  const apiKey = PropertiesService.getScriptProperties().getProperty(keyName);
  
  if (!apiKey) {
    return `${provider} API key not found. Please set the API key in the AI Chatbot menu.`;
  }

  // Add user message to conversation history
  conversationHistory.push({ role: "user", content: message });

  // Prepare the API call based on the provider
  let apiUrl, headers, payload;
  
  switch(provider) {
    case 'OpenAI':
      model = 'gpt-4o-mini';
      apiUrl = 'https://api.openai.com/v1/chat/completions';
      headers = {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      };
      payload = JSON.stringify({
        model: model,
        messages: conversationHistory,
        temperature: 0.7
      });
      break;
    case 'Groq':
      model= 'llama-3.1-70b-versatile';
      apiUrl = 'https://api.groq.com/openai/v1/chat/completions';
      headers = {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      };
      payload = JSON.stringify({
        model: model,
        messages: conversationHistory,
        temperature: 0.7,
        max_tokens: 2048,
        top_p: 0.9,
        stream: false
      });
      break;
    case 'Together':
      model = 'meta-llama/Meta-Llama-3.1-70B-Instruct-Turbo';
      apiUrl = 'https://api.together.ai/chat/completions';
      headers = {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      };
      payload = JSON.stringify({
        model: model,
        messages: conversationHistory,
        max_tokens: 2048,
        temperature: 0.7,
        top_p: 0.8,
        stop: []
      });
      break;
    case 'Google':
      model = 'gemini-1.5-flash-002';
      apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      headers = {
        'Content-Type': 'application/json'
      };
      payload = JSON.stringify({
        contents: conversationHistory.map(msg => ({
          role: msg.role === "assistant" ? "model" : msg.role,
          parts: [{ text: msg.content }]
        })),
        generationConfig: {
          temperature: 0.7,
          topK: 40,
          topP: 0.95,
          maxOutputTokens: 8192
        }
      });
      break;
    case 'Anthropic':
      model = 'claude-3-5-sonnet-20240620';
      apiUrl = 'https://api.anthropic.com/v1/messages';
      headers = {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'Content-Type': 'application/json'
      };
      payload = JSON.stringify({
        model: model,
        max_tokens: 2048,
        temperature: 0.7,
        messages: conversationHistory
      });
      break;
    case 'Hyperbolic':
      model = 'NousResearch/Hermes-3-Llama-3.1-70B';
      apiUrl = 'https://api.hyperbolic.xyz/v1/chat/completions';
      headers = {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      };
      payload = JSON.stringify({
        messages: conversationHistory,
        model: model,
        max_tokens: 2048,
        temperature: 0.7,
        top_p: 0.9,
        stream: false
      });
      break;
    case 'Mistral':
      model = 'mistral-large-latest';
      apiUrl = 'https://api.mistral.ai/v1/chat/completions';
      headers = {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      };
      payload = JSON.stringify({
        model: model,
        messages: conversationHistory,
        temperature: 0.7,
        max_tokens: 2048,
        top_p: 1,
        stream: false,
        safe_prompt: false
      });
      break;
    case 'Cerebras':
      model = 'llama3.1-70b';
      apiUrl = 'https://api.cerebras.ai/v1/chat/completions';
      headers = {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      };
      payload = JSON.stringify({
        model: model,
        messages: conversationHistory,
        temperature: 0.7,
        max_tokens: -1,
        seed: 0,
        top_p: 1,
        stream: false
      });
      break;
    case 'SambaNova':
      model = 'Meta-Llama-3.1-405B-Instruct';
      apiUrl = 'https://api.sambanova.ai/v1/chat/completions';
      headers = {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      };
      payload = JSON.stringify({
        model: model,
        messages: [
          { role: "system", content: "You are a helpful assistant" },
          ...conversationHistory
        ],
        temperature: 0.7,
        stream: false
      });
      break;
    default:
      return 'Unsupported provider';
  }

  try {
    const response = UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      headers: headers,
      payload: payload
    });

    const result = JSON.parse(response.getContentText());
    let aiResponse;

    // Extract the AI response based on the provider's response format
    switch(provider) {
      case 'Google':
        aiResponse = result.candidates?.[0]?.content?.parts?.[0]?.text || 'No response from Google';
        break;
      // case 'HuggingFace':
      //   aiResponse = result?.[0]?.generated_text || 'No response from HuggingFace';
      //   break;
      case 'Anthropic':
        aiResponse = result.content?.[0]?.text || 'No response from Anthropic';
        break;
      default:
        aiResponse = result.choices?.[0]?.message?.content || 'No response from provider';
    }

    aiResponse = aiResponse.trim();
    
    // Add AI response to conversation history
    conversationHistory.push({ role: "assistant", content: aiResponse });

    return { response: aiResponse, history: conversationHistory, model: model };
  } catch (error) {
    return { response: `Error: ${error.message || error.toString()}`, model: model };
  }
}

function getAIResponse(provider, message) {
  return chatbot(provider, message).response;
}

function getSheetValuesFromServer(input) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getRange(input);
    var values = range.getValues();
    
    // If it's a single cell, return the value directly
    if (values.length === 1 && values[0].length === 1) {
      return values[0][0];
    }
    
    // Otherwise, return the array of values
    return values;
  } catch (error) {
    console.error('Error in getSheetValuesFromServer:', error);
    return null;
  }
}

function exportMessageToSheet(message, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chat Export") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Chat Export");
    
    var lastRow = Math.max(sheet.getLastRow(), 1);
    
    if (lastRow === 1 && sheet.getRange(1, 1).getValue() === "") {
      sheet.getRange(1, 1).setValue("Timestamp");
      sheet.getRange(1, 2).setValue("Role");
      sheet.getRange(1, 3).setValue("Message");
      lastRow = 1;
    }
    
    // Remove any "You:" or "AI:" prefixes from the message
    message = message.replace(/^(?:You|AI):\s*/, '').trim();
    
    sheet.getRange(lastRow + 1, 1).setValue(new Date());
    sheet.getRange(lastRow + 1, 2).setValue(role);
    sheet.getRange(lastRow + 1, 3).setValue(message);
    
    sheet.autoResizeColumns(1, 3);
    
    return true;
  } catch (error) {
    console.error('Error in exportMessageToSheet:', error);
    return false;
  }
}

function exportEntireChatToSheet(chatText) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chat Export") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Chat Export");
    
    var lastRow = Math.max(sheet.getLastRow(), 1);
    
    if (lastRow === 1 && sheet.getRange(1, 1).getValue() === "") {
      sheet.getRange(1, 1).setValue("Timestamp");
      sheet.getRange(1, 2).setValue("Role");
      sheet.getRange(1, 3).setValue("Message");
      lastRow = 1;
    }
    
    var messages = chatText.split('\n');
    
    var data = messages
      .filter(function(message) {
        // Filter out empty lines and lines that only contain "Export"
        return message.trim() !== "" && message.trim() !== "Export";
      })
      .map(function(message) {
        var role = message.startsWith("You:") ? "You" : "AI";
        // Remove "You:" or "AI:" prefixes and trim the message
        var cleanMessage = message.replace(/^(?:You|AI):\s*/, '').trim();
        return [new Date(), role, cleanMessage];
      });
    
    if (data.length > 0) {
      sheet.getRange(lastRow + 1, 1, data.length, 3).setValues(data);
    }
    
    sheet.autoResizeColumns(1, 3);
    
    return true;
  } catch (error) {
    console.error('Error in exportEntireChatToSheet:', error);
    return false;
  }
}

function logChatMessage(message, role, provider, model) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Chat Log");
    
    if (!sheet) {
      sheet = ss.insertSheet("Chat Log");
      sheet.getRange(1, 1).setValue("Timestamp");
      sheet.getRange(1, 2).setValue("Role");
      sheet.getRange(1, 3).setValue("Message");
      sheet.getRange(1, 4).setValue("Provider");
      sheet.getRange(1, 5).setValue("Model");
    }
    
    var lastRow = Math.max(sheet.getLastRow(), 1);
    
    sheet.getRange(lastRow + 1, 1).setValue(new Date());
    sheet.getRange(lastRow + 1, 2).setValue(role);
    sheet.getRange(lastRow + 1, 3).setValue(message);
    sheet.getRange(lastRow + 1, 4).setValue(provider || "");
    sheet.getRange(lastRow + 1, 5).setValue(model || "");
    
    sheet.autoResizeColumns(1, 5);
    
    return true;
  } catch (error) {
    console.error('Error in logChatMessage:', error);
    return false;
  }
}

function doGet() {
		  return HtmlService.createHtmlOutputFromFile('Index')
			.setTitle('AI Chat Interface')
			.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
		}

/**
 * Generic function to handle API key operations.
 */
function handleApiKey(service, operation, value = null) {
  const ui = SpreadsheetApp.getUi();
  const keyName = formatKeyName(service);

  switch (operation) {
    case 'set':
      if (value && value.trim() !== '') {
        PropertiesService.getScriptProperties().setProperty(keyName, value.trim());
        ui.alert('API Key Saved', `Your ${service} API key has been saved successfully.`, ui.ButtonSet.OK);
      } else {
        ui.alert('Error', 'API key cannot be empty. Please try again.', ui.ButtonSet.OK);
      }
      break;
    case 'view':
      const apiKey = PropertiesService.getScriptProperties().getProperty(keyName) || 'Not set';
      ui.alert('Current API Key', `Your current ${service} API key is: ${apiKey}`, ui.ButtonSet.OK);
      break;
    case 'clear':
      PropertiesService.getScriptProperties().deleteProperty(keyName);
      ui.alert('API Key Cleared', `The ${service} API key has been cleared.`, ui.ButtonSet.OK);
      break;
  }
}

// OpenAI Functions
function setOpenAIApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Enter OpenAI API Key', 'Please enter your OpenAI API key:', ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    handleApiKey('OpenAI', 'set', result.getResponseText());
  }
}

function viewOpenAIApiKey() {
  handleApiKey('OpenAI', 'view');
}

function clearOpenAIApiKey() {
  handleApiKey('OpenAI', 'clear');
}

// Groq Functions
function setGroqApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set Groq API Key', 'Please enter your Groq API key:', ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    handleApiKey('Groq', 'set', result.getResponseText());
  }
}

function viewGroqApiKey() {
  handleApiKey('Groq', 'view');
}

function clearGroqApiKey() {
  handleApiKey('Groq', 'clear');
}

// Together Functions
function setTogetherApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set Together API Key', 'Please enter your Together API key:', ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    handleApiKey('Together', 'set', result.getResponseText());
  }
}

function viewTogetherApiKey() {
  handleApiKey('Together', 'view');
}

function clearTogetherApiKey() {
  handleApiKey('Together', 'clear');
}

// Google Functions
function setGoogleApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set Google API Key', 'Please enter your Google API key:', ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    handleApiKey('Google', 'set', result.getResponseText());
  }
}

function viewGoogleApiKey() {
  handleApiKey('Google', 'view');
}

function clearGoogleApiKey() {
  handleApiKey('Google', 'clear');
}

// Antrhopic Functions
function setAnthropicApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set Anthropic API Key', 'Please enter your Anthropic API key:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() == ui.Button.OK) {
    handleApiKey('Anthropic', 'set', result.getResponseText());
  }
}

function viewAnthropicApiKey() {
  handleApiKey('Anthropic', 'view');
}

function clearAnthropicApiKey() {
  handleApiKey('Anthropic', 'clear');
}

// Hyperbolic Functions
function setHyperbolicApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set Hyperbolic API Key', 'Please enter your Hyperbolic API key:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() == ui.Button.OK) {
    handleApiKey('Hyperbolic', 'set', result.getResponseText());
  }
}

function viewHyperbolicApiKey() {
  handleApiKey('Hyperbolic', 'view');
}

function clearHyperbolicApiKey() {
  handleApiKey('Hyperbolic', 'clear');
}

// Mistral Functions
function setMistralApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set Mistral API Key', 'Please enter your Mistral API key:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() == ui.Button.OK) {
    handleApiKey('Mistral', 'set', result.getResponseText());
  }
}

function viewMistralApiKey() {
  handleApiKey('Mistral', 'view');
}

function clearMistralApiKey() {
  handleApiKey('Mistral', 'clear');
}

// Cerebras Functions
function setCerebrasApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set Cerebras API Key', 'Please enter your Cerebras API key:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() == ui.Button.OK) {
    handleApiKey('Cerebras', 'set', result.getResponseText());
  }
}

function viewCerebrasApiKey() {
  handleApiKey('Cerebras', 'view');
}

function clearCerebrasApiKey() {
  handleApiKey('Cerebras', 'clear');
}

// SambaNova Functions
function setSambaNovaApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set SambaNova API Key', 'Please enter your SambaNova API key:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() == ui.Button.OK) {
    handleApiKey('SambaNova', 'set', result.getResponseText());
  }
}

function viewSambanovaApiKey() {
  handleApiKey('SambaNova', 'view');
}

function clearSambanovaApiKey() {
  handleApiKey('SambaNova', 'clear');
}
