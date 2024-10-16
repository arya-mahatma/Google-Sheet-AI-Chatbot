# AI Chatbot for Google Sheets

This project implements an AI-powered chatbot sidebar within Google Sheets, allowing users to interact with various AI models directly alongside their spreadsheet environment. The chatbot provides a convenient way to access AI capabilities while working in Google Sheets, though it does not directly manipulate spreadsheet data.

## Key Features

- Integration with multiple AI providers:
  - OpenAI
  - Groq
  - Together
  - Google (Gemini)
  - Anthropic (Claude)
  - Hyperbolic
  - Mistral
  - Cerebras
  - SambaNova
- Easy-to-use sidebar interface for chat interactions
- Support for voice input and text-to-speech output
- Option to export chat conversations to a separate sheet
- Functionality to reference spreadsheet data in conversations
- Secure API key management for each provider

## Setup Instructions

1. Open your Google Sheet.
2. Go to Extensions > Apps Script.
3. Create two new files:
   - Name one `sidebar.html` and paste the contents of the provided HTML file.
   - Name the other `Code.gs` and paste the contents of the provided JavaScript file.
4. Save the project and close the Apps Script editor.
5. Refresh your Google Sheet.
6. You should now see an "AI Chatbot" menu in your Google Sheets toolbar.

## Usage

1. **Important**: Before using any AI provider, you must first set up your API key:
   - Click on "AI Chatbot" in the menu.
   - Select the submenu for the provider you want to use (e.g., "OpenAI").
   - Click "Set API Key" and enter your API key when prompted.
   - Repeat this process for each provider you intend to use.

2. Once your API key(s) are set up, click on "AI Chatbot" > "Start AI Chatbot" to open the chatbot sidebar.
3. Select your preferred AI provider from the dropdown menu.
4. Enter your message in the text area and click "Send" or press Enter.
5. The AI's response will appear in the chat window.

## Important Security Notes

1. **API Key Protection**: Your API keys are stored in the script's properties and are not visible in the sheet itself. However, anyone with edit access to the sheet can potentially access these keys through the Apps Script project.

2. **Data Privacy**: Be cautious about sharing sheets containing sensitive conversations or data. Remember that exported chats and logs are visible to anyone with access to the sheet.

3. **Usage Limits**: Be aware of the usage limits and costs associated with each AI provider. This script does not implement rate limiting or usage tracking.

4. **Content Filtering**: This chatbot does not implement content filtering. Be mindful of the inputs provided and the potential outputs from the AI models.

## Troubleshooting

If you encounter issues:

1. Check that your API keys are correctly set and have the necessary permissions.
2. Ensure you have a stable internet connection.
3. Check the Apps Script execution log for any error messages.

## Contributing

This project is open for contributions. Feel free to fork the repository, make improvements, and submit pull requests.

## Disclaimer

This project is not officially associated with or endorsed by any of the AI providers mentioned. Use it at your own risk and in compliance with each provider's terms of service.
