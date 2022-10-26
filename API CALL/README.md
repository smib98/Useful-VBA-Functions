# VBA-API-Call

A simple demonstration of how to make an API Call to Open AI's GPT-3 from Excel using VBA. I have included two files in this example; the original curl API and the modified VBA API call.

## Before you start
• Make sure you already have an OpenAI account before you start
<br>
```
Sign up for OpenAI using the link below. Signing up usually entitles you to free credits
https://beta.openai.com/signup
```
• Get your OpenAI Key by clicking **"Personal"**, **"View API Keys"**, **"Create new secret key"**

## How To Use

1. Open the VBA editor in Excel (Alt + F11)
2. Copy the code from the file "VBA_API" into "This Workbook" in the VBA editor
3. In the VBA editor, click **Tools**, **References**, and select **"Microsoft WinHTTP Services"**
<br>`RECOMENDED - Enable "Microsoft XML" to help parse the response from the API`
4. Replace the text "OPENAI_API_KEY" in the code with your OpenAI Key
5. Change the text "Prompt goes here" with the prompt you want to send to the API
6. DONE! The response will be stored as a JSON string in the variable **"AIRESPONSE"**

## Useful Tips
• When entering your prompt / prompt parameters, use double quotes to avoid escaping the string<br>
• You can use **\n** in your prompt to send a new line
