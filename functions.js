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

// @ts-ignore
if (window.CustomFunctions) {
    // @ts-ignore
    CustomFunctions.associate("DEBUGLOG", DEBUGLOG);
    // @ts-ignore
    CustomFunctions.associate("TESTGET", TESTGET);
}
