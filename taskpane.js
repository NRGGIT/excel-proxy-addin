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
    // @ts-ignore
    window.Excel.run(function (context) {
        var settings = context.workbook.settings;
        var settingNames = ["knowledge_model_id", "api_key", "default_extension", "default_model_alias", "max_tokens", "temperature"];
        settingNames.forEach(function(name) {
            // @ts-ignore
            var value = document.getElementById(name).value;
            settings.add(name, value);
        });

        return context.sync().then(function() {
            var status = document.getElementById("status");
            status.textContent = "Settings saved.";
            setTimeout(function() {
                status.textContent = "";
            }, 3000);
        });
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
