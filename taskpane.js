Office.onReady(function () {
    // @ts-ignore
    if (window.Excel) {
        // @ts-ignore
        window.Excel.run(function (context) {
            var settings = context.workbook.settings;
            settings.onSettingsChanged.add(function (eventArgs) {
                var logDiv = document.getElementById("log");
                var newLog = document.createElement("div");
                newLog.textContent = eventArgs.value;
                logDiv.appendChild(newLog);
            });
            return context.sync();
        });
    }
});
