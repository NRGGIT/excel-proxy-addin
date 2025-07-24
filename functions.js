(function () {
    Office.onReady(function() {
        // @ts-ignore
        if (window.Excel) {
            // @ts-ignore
            window.Excel.run(function (context) {
                // @ts-ignore
                if (window.CustomFunctions) {
                    // @ts-ignore
                    window.CustomFunctions.associate("TESTGET", TESTGET);
                }
                return context.sync();
            });
        }
    });

    /**
     * Gets the text from a URL.
     * @customfunction
     * @param {string} url The URL to fetch.
     * @returns {string} The text from the URL.
     */
    function log(message) {
        // @ts-ignore
        if (window.Excel) {
            // @ts-ignore
            window.Excel.run(function (context) {
                context.workbook.settings.add("log", message);
                return context.sync();
            });
        }
    }

    function TESTGET(url) {
        log("TESTGET called with URL: " + url);
        const proxyUrl = "https://api.allorigins.win/raw?url=";
        return new Promise(function (resolve, reject) {
            fetch(proxyUrl + encodeURIComponent(url))
                .then(function (response) {
                    log("Fetch response received. Status: " + response.status);
                    if (response.ok) {
                        return response.text();
                    } else {
                        log("Error fetching URL: " + response.statusText);
                        reject(new Error("Error fetching URL: " + response.statusText));
                    }
                })
                .then(function (text) {
                    log("Response text received. Length: " + text.length);
                    resolve(text);
                })
                .catch(function (error) {
                    log("Fetch error: " + error.message);
                    reject(error);
                });
        });
    }
})();
