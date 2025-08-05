/**
 * Makes a POST request through the local server proxy
 * @param {string} url The target URL
 * @param {object} options The fetch options
 * @returns {Promise} The fetch promise
 */
function makePostRequest(url, options) {
    return new Promise(function (resolve, reject) {
        // Use local server as proxy
        const proxyUrl = "http://localhost:8000/proxy";
        
        const proxyRequest = {
            url: url,
            method: "POST",
            headers: options.headers,
            body: options.body
        };

        DEBUGLOG("Making proxy request to: " + url);
        DEBUGLOG("Proxy request data: " + JSON.stringify(proxyRequest));

        fetch(proxyUrl, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(proxyRequest)
        })
        .then(function(response) {
            DEBUGLOG("Proxy response status: " + response.status);
            if (response.ok) {
                return response;
            } else {
                return response.text().then(function(text) {
                    throw new Error("Proxy error: " + response.status + " " + text);
                });
            }
        })
        .then(function(response) {
            return response.json();
        })
        .then(function(data) {
            DEBUGLOG("Proxy response data: " + JSON.stringify(data));
            resolve(data);
        })
        .catch(function(error) {
            DEBUGLOG("Proxy error: " + error.message);
            reject(error);
        });
    });
}

/**
 * Tests API connectivity using the local server
 * @customfunction
 * @returns {string} API connectivity status
 */
function TESTAPICONNECTIVITY() {
    return new Promise(function (resolve, reject) {
        fetch("http://localhost:8000/test-api")
            .then(function(response) {
                return response.json();
            })
            .then(function(data) {
                DEBUGLOG("API connectivity test: " + JSON.stringify(data));
                resolve(data.message);
            })
            .catch(function(error) {
                DEBUGLOG("API connectivity error: " + error.message);
                resolve("Error: " + error.message);
            });
    });
}

/**
 * Validates API key by making a test request
 * @customfunction
 * @param {string} api_key The API key to validate
 * @returns {string} Validation result
 */
function VALIDATEAPIKEY(api_key) {
    return new Promise(function (resolve, reject) {
        const url = "https://constructor.app/api/platform-kmapi/alive";
        const headers = {
            "X-KM-AccessKey": "Bearer " + api_key,
            "Content-Type": "application/json",
            "Accept": "application/json"
        };

        const proxyRequest = {
            url: url,
            method: "GET",
            headers: headers,
            body: ""
        };

        DEBUGLOG("Validating API key...");
        DEBUGLOG("Validation request: " + JSON.stringify(proxyRequest));

        fetch("http://localhost:8000/proxy", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(proxyRequest)
        })
        .then(function(response) {
            DEBUGLOG("Validation response status: " + response.status);
            return response.text();
        })
        .then(function(text) {
            DEBUGLOG("Validation response: " + text);
            if (text.includes("401") || text.includes("Unauthorized")) {
                resolve("❌ Invalid API key");
            } else {
                resolve("✅ API key is valid");
            }
        })
        .catch(function(error) {
            DEBUGLOG("Validation error: " + error.message);
            resolve("❌ Validation error: " + error.message);
        });
    });
}

/**
 * Simple test function that doesn't require Excel context.
 * @customfunction SIMPLETEST
 * @param {string} message The message to return.
 * @returns {string} The message with a prefix.
 */
function SIMPLETEST(message) {
    return "Test: " + message;
}

/**
 * Writes a message to cell A1 for debugging.
 * @customfunction DEBUGLOG
 * @param {string} message The message to write.
 * @returns {string} A confirmation message.
 */
function DEBUGLOG(message) {
    try {
        // @ts-ignore
        if (window.Excel) {
            // @ts-ignore
            window.Excel.run(function (context) {
                var sheet = context.workbook.worksheets.getActiveWorksheet();
                var range = sheet.getRange("A1");
                range.values = [[message]];
                return context.sync();
            }).catch(function (error) {
                console.log("Excel Error: " + error);
                // Return error message instead of trying to write to Excel again
                return "Error: " + error.message;
            });
        } else {
            console.log("DEBUGLOG: " + message);
        }
    } catch (error) {
        console.log("DEBUGLOG Error: " + error);
        return "Error: " + error.message;
    }
    return "Logged: " + message;
}

/**
 * Gets the text from a URL.
 * @customfunction
 * @param {string} url The URL to fetch.
 * @returns {string} The text from the URL.
 */
function TESTGET(url) {
    DEBUGLOG("TESTGET called with URL: " + url);
    const proxyUrl = "https://api.allorigins.win/raw?url=";
    return new Promise(function (resolve, reject) {
        fetch(proxyUrl + encodeURIComponent(url))
            .then(function (response) {
                DEBUGLOG("Fetch response received. Status: " + response.status);
                if (response.ok) {
                    return response.text();
                } else {
                    DEBUGLOG("Error fetching URL: " + response.statusText);
                    reject(new Error("Error fetching URL: " + response.statusText));
                }
            })
            .then(function (text) {
                DEBUGLOG("Response text received. Length: " + text.length);
                resolve(text);
            })
            .catch(function (error) {
                DEBUGLOG("Fetch error: " + error.message);
                reject(error);
            });
    });
}



/**
 * Gets KMAPI settings from server
 * @returns {Promise<object>} Settings object or null if not found
 */
function getKMAPISettings() {
    return fetch("http://localhost:8000/settings", {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            action: "get"
        })
    })
    .then(function(response) {
        return response.json();
    })
    .then(function(data) {
        if (data.status === "success" && data.settings && data.settings.knowledge_model_id && data.settings.api_key) {
            console.log("Found settings from server:", data.settings);
            return data.settings;
        } else {
            console.log("No valid settings found on server:", data);
            return null;
        }
    })
    .catch(function(error) {
        console.log("Error getting settings from server:", error);
        return null;
    });
}

/**
 * Calls the KMAPI to get a chat completion.
 * @customfunction
 * @param {string} userMsg The user's message.
 * @param {string} [systemMsg] The system message.
 * @param {string} [model] The model alias.
 * @param {string} [extension] The extension.
 * @returns {string} The completion from the API.
 */
function KMAPI(userMsg, systemMsg, model, extension) {
    console.log("KMAPI function called with:", { userMsg, systemMsg, model, extension });
    
    return new Promise(function (resolve, reject) {
        try {
            // Get settings from server
            getKMAPISettings().then(function(settings) {
                console.log("Settings retrieved:", settings);
                
                if (!settings || !settings.knowledge_model_id || !settings.api_key) {
                    var errorMsg = "Please save settings in the add-in task pane first. Click the Settings button and enter your API credentials.";
                    console.log("Settings error:", errorMsg);
                    reject(new Error(errorMsg));
                    return;
                }

                if (!userMsg) {
                    var errorMsg = "userMsg parameter is required";
                    console.log("Parameter error:", errorMsg);
                    reject(new Error(errorMsg));
                    return;
                }

                var knowledge_model_id = settings.knowledge_model_id;
                var api_key = settings.api_key;
                var ext = extension || settings.default_extension || "direct_llm";
                var model_alias = model || settings.default_model_alias || "gpt-4o-2024-08-06";
                var max_tokens = settings.max_tokens || 2048;
                var temp = settings.temperature || 0.7;

                console.log("Using settings:", { knowledge_model_id, api_key: api_key.substring(0, 10) + "...", ext, model_alias, max_tokens, temp });

                // Fixed API endpoint structure
                var url = "https://constructor.app/api/platform-kmapi/v1/knowledge-models/" + knowledge_model_id + "/chat/completions/" + ext;
                
                // Fixed headers - added "Bearer " prefix
                var headers = {
                    "X-KM-AccessKey": "Bearer " + api_key,
                    "Content-Type": "application/json",
                    "Accept": "application/json"
                };

                var messages = [];
                if (systemMsg) {
                    messages.push({ role: "system", content: [{ type: "text", text: systemMsg }] });
                }
                messages.push({ role: "user", content: [{ type: "text", text: userMsg }] });

                var body = {
                    model: model_alias,
                    messages: messages,
                    response_format: { type: "text", json_schema: {} },
                    temperature: parseFloat(temp),
                    max_completion_tokens: parseInt(max_tokens),
                    top_p: 1,
                    frequency_penalty: 0,
                    presence_penalty: 0
                };

                DEBUGLOG("Making KMAPI request to: " + url);
                DEBUGLOG("Headers: " + JSON.stringify(headers));
                DEBUGLOG("Body: " + JSON.stringify(body));

                makePostRequest(url, {
                    headers: headers,
                    body: JSON.stringify(body)
                })
                .then(function(json) {
                    DEBUGLOG("API Response: " + JSON.stringify(json));
                    if (json.choices && json.choices.length > 0 && json.choices[0].message) {
                        resolve(json.choices[0].message.content);
                    } else {
                        reject(new Error("Invalid response from API: " + JSON.stringify(json)));
                    }
                })
                .catch(function(error) {
                    DEBUGLOG("KMAPI Error: " + error.message);
                    reject(error);
                });
            }).catch(function(error) {
                console.log("Error getting settings:", error);
                reject(error);
            });
        } catch (error) {
            console.log("KMAPI function error:", error);
            reject(error);
        }
    });
}

/**
 * Calls the KMAPI with all parameters provided directly.
 * @customfunction
 * @param {string} knowledge_model_id The Knowledge Model ID.
 * @param {string} api_key The API Key.
 * @param {string} userMsg The user's message.
 * @param {string} [systemMsg] The system message.
 * @param {string} [model] The model alias.
 * @param {string} [extension] The extension.
 * @param {number} [max_tokens] The max tokens.
 * @param {number} [temperature] The temperature.
 * @returns {string} The completion from the API.
 */
function KMAPITEST(knowledge_model_id, api_key, userMsg, systemMsg, model, extension, max_tokens, temperature) {
    return new Promise(function (resolve, reject) {
        var ext = extension || "direct_llm";
        var model_alias = model || "gpt-4o-2024-08-06";
        var max_tokens_val = max_tokens || 2048;
        var temp_val = temperature || 0.7;

        // Fixed API endpoint structure
        var url = "https://constructor.app/api/platform-kmapi/v1/knowledge-models/" + knowledge_model_id + "/chat/completions/" + ext;
        
        // Fixed headers - added "Bearer " prefix
        var headers = {
            "X-KM-AccessKey": "Bearer " + api_key,
            "Content-Type": "application/json",
            "Accept": "application/json"
        };

        var messages = [];
        if (systemMsg) {
            messages.push({ role: "system", content: [{ type: "text", text: systemMsg }] });
        }
        messages.push({ role: "user", content: [{ type: "text", text: userMsg }] });

        var body = {
            model: model_alias,
            messages: messages,
            response_format: { type: "text", json_schema: {} },
            temperature: parseFloat(temp_val),
            max_completion_tokens: parseInt(max_tokens_val),
            top_p: 1,
            frequency_penalty: 0,
            presence_penalty: 0
        };

        DEBUGLOG("KMAPITEST - URL: " + url);
        DEBUGLOG("KMAPITEST - Headers: " + JSON.stringify(headers));
        DEBUGLOG("KMAPITEST - Body: " + JSON.stringify(body));

        makePostRequest(url, {
            headers: headers,
            body: JSON.stringify(body)
        })
        .then(function(json) {
            DEBUGLOG("KMAPITEST Response: " + JSON.stringify(json));
            if (json.choices && json.choices.length > 0 && json.choices[0].message) {
                resolve(json.choices[0].message.content);
            } else {
                reject(new Error("Invalid response from API: " + JSON.stringify(json)));
            }
        })
        .catch(function(error) {
            DEBUGLOG("KMAPITEST Error: " + error.message);
            reject(error);
        });
    });
}

/**
 * Tests the API connection and returns diagnostic information.
 * @customfunction
 * @param {string} knowledge_model_id The Knowledge Model ID.
 * @param {string} api_key The API Key.
 * @returns {string} Diagnostic information about the API connection.
 */
function TESTAPI(knowledge_model_id, api_key) {
    return new Promise(function (resolve, reject) {
        var url = "https://constructor.app/api/platform-kmapi/v1/knowledge-models/" + knowledge_model_id + "/chat/completions/direct_llm";
        var headers = {
            "X-KM-AccessKey": "Bearer " + api_key,
            "Content-Type": "application/json",
            "Accept": "application/json"
        };

        var body = {
            model: "gpt-4o-2024-08-06",
            messages: [{ role: "user", content: [{ type: "text", text: "Hello" }] }],
            response_format: { type: "text", json_schema: {} },
            temperature: 0.7,
            max_completion_tokens: 10,
            top_p: 1,
            frequency_penalty: 0,
            presence_penalty: 0
        };

        DEBUGLOG("TESTAPI - Testing connection to: " + url);
        DEBUGLOG("TESTAPI - Headers: " + JSON.stringify(headers));

        makePostRequest(url, {
            headers: headers,
            body: JSON.stringify(body)
        })
        .then(function(json) {
            DEBUGLOG("TESTAPI - Response body: " + JSON.stringify(json));
            
            // Extract the assistant's content from the response
            if (json.choices && json.choices.length > 0 && json.choices[0].message) {
                var content = json.choices[0].message.content;
                resolve("✅ API connection successful! Assistant response: " + content);
            } else {
                resolve("✅ API connection successful! Response: " + JSON.stringify(json));
            }
        })
        .catch(function(error) {
            DEBUGLOG("TESTAPI - Error: " + error.message);
            resolve("❌ API connection failed: " + error.message);
        });
    });
}

/**
 * Simple test to check if server settings work in custom functions
 * @customfunction
 * @returns {string} Test result
 */
function TESTGLOBALS() {
    console.log("TESTGLOBALS called");
    
    return fetch("http://localhost:8000/settings", {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            action: "get"
        })
    })
    .then(function(response) {
        return response.json();
    })
    .then(function(data) {
        if (data.status === "success" && data.settings && data.settings.knowledge_model_id && data.settings.api_key) {
            return "✅ Server settings found: " + JSON.stringify(data.settings);
        } else {
            return "❌ No settings found on server. Please save settings in the task pane first.";
        }
    })
    .catch(function(error) {
        return "❌ Error getting settings from server: " + error.message;
    });
}

// @ts-ignore
if (window.CustomFunctions) {
    // @ts-ignore
    CustomFunctions.associate("SIMPLETEST", SIMPLETEST);
    // @ts-ignore
    CustomFunctions.associate("DEBUGLOG", DEBUGLOG);
    // @ts-ignore
    CustomFunctions.associate("TESTGET", TESTGET);
    // @ts-ignore
    CustomFunctions.associate("KMAPI", KMAPI);
    // @ts-ignore
    CustomFunctions.associate("KMAPITEST", KMAPITEST);
    // @ts-ignore
    CustomFunctions.associate("TESTAPI", TESTAPI);
    // @ts-ignore
    CustomFunctions.associate("TESTAPICONNECTIVITY", TESTAPICONNECTIVITY);
    // @ts-ignore
    CustomFunctions.associate("VALIDATEAPIKEY", VALIDATEAPIKEY);
    // @ts-ignore
    CustomFunctions.associate("TESTGLOBALS", TESTGLOBALS);
}
