Office.onReady(function () {
    // @ts-ignore
    if (window.Excel) {
        loadSettings();
        document.getElementById("save_settings").onclick = saveSettings;
    }
});

function loadSettings() {
    // @ts-ignore
    window.Excel.run(function (context) {
        var settings = context.workbook.settings;
        var settingNames = ["knowledge_model_id", "api_key", "default_extension", "default_model_alias", "max_tokens", "temperature"];
        var settingObjects = settingNames.map(function(name) {
            return settings.getItem(name);
        });

        return context.sync().then(function() {
            settingObjects.forEach(function(setting, index) {
                if (setting.value) {
                    // @ts-ignore
                    document.getElementById(settingNames[index]).value = setting.value;
                }
            });
        });
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

function saveSettings() {
    var api_key = document.getElementById("api_key").value;
    var knowledge_model_id = document.getElementById("knowledge_model_id").value;
    var default_extension = document.getElementById("default_extension").value;
    var default_model_alias = document.getElementById("default_model_alias").value;
    var max_tokens = document.getElementById("max_tokens").value;
    var temperature = document.getElementById("temperature").value;
    
    // Show saving status
    var status = document.getElementById("status");
    status.textContent = "Saving settings to server and validating API key...";
    
    // Save the settings to Excel
    // @ts-ignore
    window.Excel.run(function (context) {
        var settings = context.workbook.settings;
        var settingNames = ["knowledge_model_id", "api_key", "default_extension", "default_model_alias", "max_tokens", "temperature"];
        var settingValues = [knowledge_model_id, api_key, default_extension, default_model_alias, max_tokens, temperature];
        
        settingNames.forEach(function(name, index) {
            settings.add(name, settingValues[index]);
        });

        return context.sync().then(function() {
            // Save settings to server
            var serverSettings = {
                knowledge_model_id: knowledge_model_id,
                api_key: api_key,
                default_extension: default_extension,
                default_model_alias: default_model_alias,
                max_tokens: parseInt(max_tokens),
                temperature: parseFloat(temperature)
            };
            
            return fetch("http://localhost:8000/settings", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    action: "save",
                    settings: serverSettings
                })
            });
        })
        .then(function(response) {
            return response.json();
        })
        .then(function(data) {
            console.log("Settings saved to server:", data);
            
            // Now validate the API key
            return validateApiKey(api_key, knowledge_model_id);
        });
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
        status.textContent = "Error saving settings: " + error.message;
        setTimeout(function() {
            status.textContent = "";
        }, 5000);
    });
}

function validateApiKey(api_key, knowledge_model_id) {
    if (!api_key) {
        var status = document.getElementById("status");
        status.textContent = "Settings saved. No API key provided for validation.";
        setTimeout(function() {
            status.textContent = "";
        }, 3000);
        return;
    }

    // Test API connectivity first
    fetch("http://localhost:8000/test-api")
        .then(function(response) {
            return response.json();
        })
        .then(function(data) {
            console.log("API connectivity test: " + JSON.stringify(data));
            
            // Now validate the API key
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

            return fetch("http://localhost:8000/proxy", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(proxyRequest)
            });
        })
        .then(function(response) {
            return response.text();
        })
        .then(function(text) {
            var status = document.getElementById("status");
            if (text.includes("401") || text.includes("Unauthorized")) {
                status.textContent = "Settings saved. ❌ API key validation failed - check your key.";
            } else {
                status.textContent = "Settings saved. ✅ API key is valid!";
            }
            setTimeout(function() {
                status.textContent = "";
            }, 5000);
        })
        .catch(function(error) {
            console.log("API validation error: " + error);
            var status = document.getElementById("status");
            status.textContent = "Settings saved. ⚠️ API validation failed: " + error.message;
            setTimeout(function() {
                status.textContent = "";
            }, 5000);
        });
}
