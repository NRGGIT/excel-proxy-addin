/**
 * Writes a message to cell A1 for debugging.
 * @customfunction DEBUGLOG
 * @param {string} message The message to write.
 * @returns {string} A confirmation message.
 */
function DEBUGLOG(message) {
    // @ts-ignore
    if (window.Excel) {
        // @ts-ignore
        window.Excel.run(function (context) {
            var sheet = context.workbook.worksheets.getActiveWorksheet();
            var range = sheet.getRange("A1");
            range.values = [[message]];
            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            // @ts-ignore
            if (window.Excel.run) {
                // @ts-ignore
                window.Excel.run(function (ctx) {
                    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1");
                    range.values = [["Error logging: " + error.message]];
                    return ctx.sync();
                });
            }
        });
    }
    return "Logged.";
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
 * Calls the KMAPI to get a chat completion.
 * @customfunction
 * @param {string} userMsg The user's message.
 * @param {string} [systemMsg] The system message.
 * @param {string} [model] The model alias.
 * @param {string} [extension] The extension.
 * @returns {string} The completion from the API.
 */
function KMAPI(userMsg, systemMsg, model, extension) {
    return new Promise(function (resolve, reject) {
        // @ts-ignore
        window.Excel.run(function (context) {
            var settings = context.workbook.settings;
            var settingNames = ["knowledge_model_id", "api_key", "default_extension", "default_model_alias", "max_tokens", "temperature"];
            var settingObjects = settingNames.map(function(name) {
                return settings.getItem(name);
            });

            return context.sync()
                .then(function() {
                    var settingsValues = {};
                    settingObjects.forEach(function(setting, index) {
                        settingsValues[settingNames[index]] = setting.value;
                    });

                    var knowledge_model_id = settingsValues["knowledge_model_id"];
                    var api_key = settingsValues["api_key"];
                    var ext = extension || settingsValues["default_extension"];
                    var model_alias = model || settingsValues["default_model_alias"];
                    var max_tokens = settingsValues["max_tokens"];
                    var temp = settingsValues["temperature"];

                    if (!knowledge_model_id || !api_key) {
                        reject(new Error("knowledge_model_id and api_key must be set in the task pane."));
                        return;
                    }

                    var targetUrl = "https://constructor.app/api/platform-kmapi/v1/knowledge-models/" + knowledge_model_id + "/chat/completions/" + ext;
                    
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

                    var proxyUrl = "https://api.allorigins.win/raw?url=" + encodeURIComponent(targetUrl);

                    fetch(proxyUrl, {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json",
                            "X-KM-AccessKey": "Bearer " + api_key
                        },
                        body: JSON.stringify(body)
                    })
                    .then(function(response) {
                        if (response.ok) {
                            return response.json();
                        } else {
                            response.text().then(function(text) {
                                reject(new Error("API Error: " + response.status + " " + text));
                            });
                        }
                    })
                    .then(function(json) {
                        if (json.choices && json.choices.length > 0) {
                            resolve(json.choices[0].message.content);
                        } else {
                            reject(new Error("Invalid response from API."));
                        }
                    })
                    .catch(function(error) {
                        reject(error);
                    });
                })
                .catch(function(error) {
                    reject(new Error("Failed to load settings: " + error.message));
                });
        });
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
        var model_alias = model || "gpt4.1-mini";
        var max_tokens_val = max_tokens || 2048;
        var temp_val = temperature || 0.7;

        var targetUrl = "https://constructor.app/api/platform-kmapi/v1/knowledge-models/" + knowledge_model_id + "/chat/completions/" + ext;
        
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

        var proxyUrl = "https://api.allorigins.win/raw?url=" + encodeURIComponent(targetUrl);

        fetch(proxyUrl, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "X-KM-AccessKey": "Bearer " + api_key
            },
            body: JSON.stringify(body)
        })
        .then(function(response) {
            if (response.ok) {
                return response.json();
            } else {
                response.text().then(function(text) {
                    reject(new Error("API Error: " + response.status + " " + text));
                });
            }
        })
        .then(function(json) {
            if (json.choices && json.choices.length > 0) {
                resolve(json.choices[0].message.content);
            } else {
                reject(new Error("Invalid response from API."));
            }
        })
        .catch(function(error) {
            reject(error);
        });
    });
}

// @ts-ignore
if (window.CustomFunctions) {
    // @ts-ignore
    CustomFunctions.associate("DEBUGLOG", DEBUGLOG);
    // @ts-ignore
    CustomFunctions.associate("TESTGET", TESTGET);
    // @ts-ignore
    CustomFunctions.associate("KMAPI", KMAPI);
    // @ts-ignore
    CustomFunctions.associate("KMAPITEST", KMAPITEST);
}
