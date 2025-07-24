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
    function TESTGET(url) {
        const proxyUrl = "https://cors-anywhere.herokuapp.com/";
        return new Promise(function (resolve, reject) {
            fetch(proxyUrl + url)
                .then(function (response) {
                    if (response.ok) {
                        return response.text();
                    } else {
                        reject(new Error("Error fetching URL: " + response.statusText));
                    }
                })
                .then(function (text) {
                    resolve(text);
                })
                .catch(function (error) {
                    reject(error);
                });
        });
    }
})();
