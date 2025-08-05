# Excel KMAPI Add-in

This Excel add-in allows you to make POST requests to the KMAPI (Knowledge Model API) directly from Excel using custom functions.

## Issues Fixed

The original add-in had several issues that prevented POST requests from working:

1. **CORS Proxy Issues**: Relied on unreliable external proxies
2. **Incorrect API Headers**: Wrong authentication header format
3. **Poor Error Handling**: Limited debugging capabilities
4. **External Dependencies**: Manifest pointed to external URLs

## Solutions Implemented

### 1. Multiple Proxy Fallbacks
The add-in now tries multiple CORS proxies in sequence:
- Direct request (no proxy)
- cors-anywhere.herokuapp.com
- thingproxy.freeboard.io
- corsproxy.io

### 2. Fixed Authentication
- Removed "Bearer " prefix from API key header
- Added proper Accept headers
- Improved error handling

### 3. Enhanced Debugging
- Added comprehensive logging to cell A1
- Better error messages
- Request/response logging

### 4. Local Development
- Updated manifest to use local file references
- Added local server for testing

## Setup Instructions

### 1. Start Local Server
```bash
python server.py
```
This starts a local server at `http://localhost:8000`

### 2. Load Add-in in Excel
1. Open Excel
2. Go to **Insert** > **My Add-ins** > **Upload My Add-in**
3. Browse to the `manifest.xml` file
4. Or use the URL: `http://localhost:8000/manifest.xml`

### 3. Configure Settings
1. Click the **Settings** button in the **KMAPI** group on the **Home** tab
2. Enter your:
   - **Knowledge Model ID**
   - **API Key**
   - **Default Extension** (default: direct_llm)
   - **Default Model Alias** (default: gpt4.1-mini)
   - **Max Tokens** (default: 2048)
   - **Temperature** (default: 0.7)
3. Click **Save Settings**

## Available Functions

### DEBUGLOG(message)
Writes debug messages to cell A1.
```
=DEBUGLOG("Testing connection")
```

### TESTGET(url)
Tests GET requests through a proxy.
```
=TESTGET("https://api.example.com/data")
```

### TESTAPI(knowledge_model_id, api_key)
Tests the API connection and returns diagnostic information.
```
=TESTAPI("your_model_id", "your_api_key")
```

### KMAPI(userMsg, [systemMsg], [model], [extension])
Makes a POST request to KMAPI using saved settings.
```
=KMAPI("Hello, how are you?")
=KMAPI("Explain quantum physics", "You are a helpful assistant", "gpt4.1-mini", "direct_llm")
```

### KMAPITEST(knowledge_model_id, api_key, userMsg, [systemMsg], [model], [extension], [max_tokens], [temperature])
Makes a POST request with all parameters provided directly.
```
=KMAPITEST("your_model_id", "your_api_key", "Hello world")
```

## Troubleshooting

### If POST requests fail:
1. Check the debug log in cell A1 using `=DEBUGLOG("test")`
2. Verify your API credentials are correct
3. Try different proxies by checking the debug output
4. Ensure your Knowledge Model ID and API Key are valid

### Common Issues:
- **CORS errors**: The add-in will automatically try multiple proxies
- **Authentication errors**: Make sure your API key is correct (without "Bearer " prefix)
- **Network errors**: Check your internet connection and firewall settings

## Development

### File Structure:
```
constr-excel/
├── manifest.xml          # Add-in configuration
├── taskpane.html        # Settings UI
├── taskpane.js          # Settings logic
├── functions.js         # Custom Excel functions
├── functions.json       # Function metadata
├── functions.html       # Function loader
├── assets/             # Icons
└── server.py           # Local development server
```

### Making Changes:
1. Edit the JavaScript files
2. Restart the local server
3. Reload the add-in in Excel

## API Endpoint Structure

The add-in makes POST requests to:
```
https://constructor.app/api/platform-kmapi/v1/knowledge-models/{knowledge_model_id}/chat/completions/{extension}
```

With headers:
```
X-KM-AccessKey: {api_key}
Content-Type: application/json
Accept: application/json
```

## Security Notes

- API keys are stored in Excel workbook settings (not secure for production)
- Consider implementing proper key management for production use
- The add-in includes multiple CORS proxies for reliability but may not work in all environments 