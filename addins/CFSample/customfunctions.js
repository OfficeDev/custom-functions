// Returns a number
function add42(num1, num2) {
    return num1 + num2 + 42;
}

// Returns a string
function getDay() {
    var d = new Date();
    var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    return days[d.getDay()];
}

// Returns a boolean
function isPrime(num) {
    if (num <= 1 || !Number.isInteger(num)) return false; // validates input
    // checks natural numbers below the square root (not optimal)
    for (var factor = Math.floor(Math.sqrt(num)); factor > 1; factor--) {
        if (num % factor === 0) return false;
    }
    return true;
}

// Computation-intensive for high inputs
function nthPrime(n) {
    var primeCount = 0;
    for (var num = 2; primeCount < n; num++) {
        if (isPrime(num)) primeCount++;
    }
    return num - 1;
}

// Simulate calling an external webservice to calculate result
function add42Promise(num1, num2) {
    return new Promise(function (resolve) {
        setTimeout(function () {
            resolve(num1 + num2 + 42);
        }, 1000);
    });
}




// Range input
function secondHighest(range) {
    var highest = range[0][0], secondHighest = range[0][0];

    for (var i = 0; i < range.length; i++) {
        for (var j = 0; j < range[i].length; j++) {
            if (range[i][j] >= highest) {
                secondHighest = highest;
                highest = range[i][j];
            }
            else if (range[i][j] >= secondHighest) {
                secondHighest = range[i][j];
            }
        }
    }

    return secondHighest;
}

// Range output
function makeArray(rows, columns) {
    var items = new Array(rows);
    for (var i = 0; i < rows; i++) {
        items[i] = new Array(columns);
        for (var j = 0; j < columns; j++) {
            items[i][j] = i + j;
        }
    }
    return items;
}



// Optional parameter
function checkOptionalParam(requiredParam, optionalParam) {
    if (optionalParam === null) {
        return "Optional parameter is NOT passed in";
    }
    else {
        return "Optional parameter is passed in " + optionalParam;
    }
}



// Volatile option
function rand(max) {
    return Math.floor(Math.random() * Math.floor(max));
}

// Streaming and cancelable option
function incrementStream(increment, caller) {
    var result = 0;

    var myInterval = setInterval(function () {
        result += increment;
        caller.setResult(result);
    }, 1000);

    caller.onCanceled = function () {
        clearInterval(myInterval);
    };
}

// requiresAddress option
function cellAddress(caller) {
    return caller.address;
}


// Write to asyncStorage
function setAsyncStorage(key, value) {
    return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
        return "Success: Item with key '" + key + "' saved to AsyncStorage.";
    }, function (error) {
        return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
    });
}

// Read from asyncStorage
function getAsyncStorage(key) {
    return OfficeRuntime.AsyncStorage.getItem(key);
}

// XMLHttpRequest API
function translate(text, langLocale) {
    return new Promise(function (resolve) {
        var xhr = new XMLHttpRequest();
        var textStr = encodeURIComponent(text);
        var localeStr = encodeURIComponent(langLocale);

        var url = "https://excelcf-demo-api.azurewebsites.net/api/translate?code=<REMOVED>&name=" +
            textStr + "&locale=" + localeStr;

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                resolve(xhr.responseText);
            }
        };

        xhr.open('GET', url, true);
        xhr.send();
    });
}

// XMLHttpRequest API
function stockPrice(ticker) {
    return new Promise(
        function (resolve) {
            let xhr = new XMLHttpRequest();
            let url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price"
            //add handler for xhr

            xhr.onreadystatechange = function () {
                if (xhr.readyState === XMLHttpRequest.DONE) {
                    //return result back to Excel

                    resolve(xhr.responseText);
                }
            };
            //make request

            xhr.open('GET', url, true);
            xhr.send();
        });
}

// XMLHttpRequest API
function stockPriceStream(ticker, caller) {

    let result = 0;

    setInterval(function () {
        let xhr = new XMLHttpRequest();
        let url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        //add handler for xhr

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                //return result back to Excel

                caller.setResult(xhr.responseText);
            }
        };

        xhr.open('GET', url, true);
        xhr.send();
    }, 1000); //milliseconds
}

// fetch API
function stockPriceHistory(ticker, days) {
    return new Promise(function (resolve) {
        var url = "https://excelcf-demo-api.azurewebsites.net/api/StockHistory?stock=" + ticker + "&days=" + days;

        fetch(url)
            .then(function (response) {
                return response.text();
            })
            .then(function (text) {
                resolve(JSON.parse(text));
            })
            .catch(function (error) {
                //will show #VALUE!
                console.log(error);
                resolve(error);
            });
    });
}

// displayWebDialog API
function displayWebDialog() {
    OfficeRuntime.displayWebDialog("https://cfsample.azurewebsites.net/home.html", {
        height: "50",
        width: "50",
        hideTitle: false,
    });
}

// WebSocket API
function bitcoinStream(increment, handler) {
    const socket = new WebSocket('wss://ws-feed.pro.coinbase.com');
    var isOpen = 0;
    var counter = increment;

    socket.onmessage = function (event) {
        counter++;
        if (counter < 1000) {
            handler.setResult("Event Counter:" + counter + " Data:" + event.data);
        }
        else {
            socket.close();
        }
    };

    socket.onclose = function () {
        isOpen = 0;
        handler.setResult("Socket closed after Event Counter:" + counter);
    };

    socket.onerror = function () {
        isOpen = 0;
        handler.setResult("Error after Event Counter:" + counter);
        socket.close();
    };

    socket.onopen = function () {
        console.log("onopen");
        socket.send(
            JSON.stringify({
                "type": "subscribe",
                "product_ids": [
                    "ETH-USD",
                    "ETH-EUR"
                ],
                "channels": [
                    "level2",
                    "heartbeat",
                    {
                        "name": "ticker",
                        "product_ids": [
                            "ETH-BTC",
                            "ETH-USD"
                        ]
                    }
                ]
            }
            ));
        isOpen = 1;
    };
}

function errorOut(how) {
	if (how === 'throw') {
		throw { prop1: "Jabberwocky thrown" };
	}
	else if (how === 'promise') {
		return Promise.reject({ prop1: "Jabberwocky rejected" });
	}
	return undefined;
}

function customErrorReturn(errorCase) {
	switch (errorCase) {
		case 1: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,	// #N/A
				"An error *case 1* was detected in customErrorReturn"
			);
			return error;
		}
		case 2: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidValue, // #VALUE!
				"An error *case 2* was detected in customErrorReturn"
			);
			return error;
		}
		case 3: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.divisionByZero, // #DIV/0!
				"This message should not appear in UI"
			);
			return error;
		}
		case 4: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidNumber, // #NUM!
				"This message should not appear in UI"
			);
			return error;
		}
		case 5: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.nullReference, // #NULL!
				"This message should not appear in UI"
			);
			return error;
		}
		case 6: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable // #N/A!
			);
			return error;
		}
		case 7: {
			var error = new CustomFunctions.Error(); // #VALUE!
			return error;
		}
		case 8: {
			var error = new CustomFunctions.Error(
				undefined, // #VALUE!
				"This message should not appear in UI"
			);
			return error;
		}
		case 9: {
			var error = new CustomFunctions.Error(
				"Customized", // #VALUE!
				"This message should not appear in UI"
			);
			return error;
		}
		case 10:{
			var error = new CustomFunctions.Error(new Error()); // #VALUE!
			return error;
		}
		case 11: {
			return new Error("This message should not appear in UI"); // #VALUE!
		}
		default: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,
				"An unknown error case was detected in customErrorReturn"
			);
			return error;
		}
	}
}

function customErrorReturnArray(errorCase) {
	switch (errorCase) {
		case 1: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,	// #N/A
				"An error *case 1* was detected in customErrorReturnArray "
			);
			return [['Hello'],[error]];
		}
		case 2: {
			var error1 = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,	// #N/A
				"An error *case 2* was detected in customErrorReturnArray "
			);
			var error2 = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidValue, // #VALUE!
				"An error *case 2* was detected in customErrorReturnArray "
			);
			return [[error1],[error2]];
		}
	}
}

CustomFunctions.associate('ADD42', add42);
CustomFunctions.associate('GET_DAY', getDay);
CustomFunctions.associate('IS_PRIME', isPrime);
CustomFunctions.associate('NTH_PRIME', nthPrime);
CustomFunctions.associate('ADD42_PROMISE', add42Promise);
CustomFunctions.associate('SECOND_HIGHEST', secondHighest);
CustomFunctions.associate('MAKE_ARRAY', makeArray);
CustomFunctions.associate('CHECK_OPTIONAL_PARAM', checkOptionalParam);
CustomFunctions.associate('RAND', rand);
CustomFunctions.associate('INCREMENT_STREAM', incrementStream);
CustomFunctions.associate('CELL_ADDRESS', cellAddress);
CustomFunctions.associate('SET_ASYNC_STORAGE', setAsyncStorage);
CustomFunctions.associate('GET_ASYNC_STORAGE', getAsyncStorage);
CustomFunctions.associate('TRANSLATE', translate);
CustomFunctions.associate('STOCK_PRICE', stockPrice);
CustomFunctions.associate('STOCK_PRICE_STREAM', stockPriceStream);
CustomFunctions.associate('STOCK_PRICE_HISTORY', stockPriceHistory);
CustomFunctions.associate('DISPLAY_WEB_DIALOG', displayWebDialog);
CustomFunctions.associate('BITCOIN_STREAM', bitcoinStream);
CustomFunctions.associate('ERROROUT', errorOut);
CustomFunctions.associate('CUSTOM_ERROR_RETURN', customErrorReturn);
CustomFunctions.associate('CUSTOM_ERROR_RETURNARRAY', customErrorReturnArray);



