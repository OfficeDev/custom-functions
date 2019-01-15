var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var functionNamespace = "MICROSOFT.OFFICE.TEST.PERF.";
function customFunctions_runAutomatedTests() {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            Excel.run(function (ctx) { return __awaiter(_this, void 0, void 0, function () {
                var inputSheetName, inputDataRangeName, range, outputSheet, radios, i, radioButton, numberOfFunctions, expectedTimeout, numberOfIterations, functionName, functionId, functionParameters, numberOfFunctionsOnEachRow, testName, expectedEventCount, outputData, i, addRangeInputParameter, addRangeInputParameter, executionTime, err_1, entireSheet, outputRange;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            inputSheetName = "Async Input";
                            inputDataRangeName = "InputTests";
                            range = ctx.workbook.worksheets.getItem(inputSheetName).getRange(inputDataRangeName);
                            ctx.load(range);
                            return [4 /*yield*/, ctx.sync()];
                        case 1:
                            _a.sent();
                            radios = document.getElementsByName('OutputSheet');
                            for (i = 0; i < radios.length; i++) {
                                radioButton = radios[i];
                                if (radioButton.checked) {
                                    outputSheet = radioButton.value;
                                    break;
                                }
                            }
                            numberOfFunctions = 0;
                            expectedTimeout = 0;
                            numberOfIterations = 0;
                            functionId = 0;
                            numberOfFunctionsOnEachRow = 0;
                            expectedEventCount = 0;
                            outputData = [];
                            outputData.push(["Test Id", "Test Name", "Avg. time (ms) per function"]);
                            logger.clear();
                            i = 1;
                            _a.label = 2;
                        case 2:
                            if (!(i < range.rowCount)) return [3 /*break*/, 11];
                            // This order must be kept in sync with the columns in the 'Input' sheet in the perf xlsx file
                            functionId = range.values[i][0];
                            functionName = functionNamespace + range.values[i][1];
                            numberOfFunctions = range.values[i][2];
                            numberOfIterations = range.values[i][3];
                            numberOfFunctionsOnEachRow = range.values[i][4];
                            expectedTimeout = range.values[i][5];
                            functionParameters = '(' + range.values[i][6] + ')';
                            expectedEventCount = range.values[i][7] + 1;
                            if (!(range.values[i][1] == "addRangeJS" || range.values[i][1] == "syncAddRangeJS")) return [3 /*break*/, 4];
                            addRangeInputParameter = range.values[i][6];
                            return [4 /*yield*/, createAddRangeInputData(ctx, addRangeInputParameter, true)];
                        case 3:
                            _a.sent();
                            _a.label = 4;
                        case 4:
                            if (!(range.values[i][1] == "addStringRangeJS" || range.values[i][1] == "syncAddStringRangeJS")) return [3 /*break*/, 6];
                            addRangeInputParameter = range.values[i][6];
                            return [4 /*yield*/, createAddRangeInputData(ctx, addRangeInputParameter, false)];
                        case 5:
                            _a.sent();
                            _a.label = 6;
                        case 6:
                            testName = range.values[i][1] + '_' + numberOfFunctions;
                            if (expectedEventCount > 1) {
                                testName += '_' + 'chainedCalls' + '_' + (expectedEventCount - 1);
                            }
                            _a.label = 7;
                        case 7:
                            _a.trys.push([7, 9, , 10]);
                            logger.comment("Running test [id = " + functionId + ", function = " + functionName + functionParameters + " ]");
                            return [4 /*yield*/, customFunctions_runAutomatedTestByTestIdHelper(ctx, range, i, true)];
                        case 8:
                            executionTime = _a.sent();
                            outputData.push([functionId, testName, executionTime]);
                            return [3 /*break*/, 10];
                        case 9:
                            err_1 = _a.sent();
                            logger.comment("Error:" + JSON.stringify(err_1));
                            outputData.push([functionId, testName, "Error:" + JSON.stringify(err_1)]);
                            return [3 /*break*/, 10];
                        case 10:
                            i++;
                            return [3 /*break*/, 2];
                        case 11:
                            logger.comment("Writting results");
                            entireSheet = ctx.workbook.worksheets.getItem(outputSheet).getRange(null);
                            entireSheet.clear(null);
                            outputRange = ctx.workbook.worksheets.getItem(outputSheet).getRange('A1:C' + range.rowCount);
                            outputRange.values = outputData;
                            return [4 /*yield*/, ctx.sync()];
                        case 12:
                            _a.sent();
                            return [2 /*return*/];
                    }
                });
            }); })
                .catch(function (ex) {
                logger.comment("Error: " + JSON.stringify(ex));
            });
            return [2 /*return*/];
        });
    });
}
function customFunctions_runAutomatedTestByTestId(testId) {
    var _this = this;
    Excel.run(function (ctx) { return __awaiter(_this, void 0, void 0, function () {
        var isAsyncExecutionElement, isAsyncExecution, isStressTestElement, isStressTest, inputSheetName, inputDataRangeName, range;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    isAsyncExecutionElement = document.getElementById("IsAsyncExecution");
                    isAsyncExecution = isAsyncExecutionElement.checked;
                    isStressTestElement = document.getElementById("IfExecuteStressTest");
                    isStressTest = isStressTestElement.checked;
                    inputSheetName = isStressTest ? (isAsyncExecution ? "Async Stress Test Input" : "Sync Stress Test Input") : (isAsyncExecution ? "Async Input" : "Sync Input");
                    inputDataRangeName = "InputTests";
                    range = ctx.workbook.worksheets.getItem(inputSheetName).getRange(inputDataRangeName);
                    ctx.load(range);
                    return [4 /*yield*/, ctx.sync()];
                case 1:
                    _a.sent();
                    customFunctions_runAutomatedTestByTestIdHelper(ctx, range, testId, isAsyncExecution);
                    return [2 /*return*/];
            }
        });
    }); })
        .catch(function (ex) {
        logger.comment("Error: " + JSON.stringify(ex));
    });
}
function customFunctions_runAutomatedTestByTestIdHelper(ctx, range, testId, isAsyncExecution) {
    return __awaiter(this, void 0, void 0, function () {
        var numberOfFunctions, expectedTimeout, numberOfIterations, functionName, functionId, functionParameters, numberOfFunctionsOnEachRow, expectedEventCount, outputData, executionTime, addRangeInputParameter, addRangeInputParameter;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    numberOfFunctions = 0;
                    expectedTimeout = 0;
                    numberOfIterations = 0;
                    functionId = 0;
                    numberOfFunctionsOnEachRow = 0;
                    expectedEventCount = 0;
                    outputData = [];
                    executionTime = 0;
                    outputData.push(["Test Id", "Test Name", "Avg. time (ms) per function"]);
                    logger.clear();
                    functionId = range.values[testId][0];
                    functionName = functionNamespace + range.values[testId][1];
                    numberOfFunctions = range.values[testId][2];
                    numberOfIterations = range.values[testId][3];
                    numberOfFunctionsOnEachRow = range.values[testId][4];
                    expectedTimeout = range.values[testId][5];
                    functionParameters = '(' + range.values[testId][6] + ')';
                    expectedEventCount = range.values[testId][7] + 1;
                    logger.comment("Running test [id = " + functionId + ", function = " + functionName + functionParameters + " ]");
                    if (!(range.values[testId][1] == "addRangeJS" || range.values[testId][1] == "syncAddRangeJS")) return [3 /*break*/, 2];
                    addRangeInputParameter = range.values[testId][6];
                    return [4 /*yield*/, createAddRangeInputData(ctx, addRangeInputParameter, true)];
                case 1:
                    _a.sent();
                    _a.label = 2;
                case 2:
                    if (!(range.values[testId][1] == "addStringRangeJS" || range.values[testId][1] == "syncAddStringRangeJS")) return [3 /*break*/, 4];
                    addRangeInputParameter = range.values[testId][6];
                    return [4 /*yield*/, createAddRangeInputData(ctx, addRangeInputParameter, false)];
                case 3:
                    _a.sent();
                    _a.label = 4;
                case 4:
                    if (!isAsyncExecution) return [3 /*break*/, 6];
                    return [4 /*yield*/, customFunctions_AsyncPerfHelper(ctx, numberOfIterations, numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters, expectedTimeout, expectedEventCount)];
                case 5:
                    executionTime = _a.sent();
                    return [3 /*break*/, 7];
                case 6:
                    customFunctions_SyncPerfHelper(ctx, numberOfIterations, numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters);
                    _a.label = 7;
                case 7: return [2 /*return*/, executionTime];
            }
        });
    });
}
function customFunctions_createTestData() {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            Excel.run(function (ctx) { return __awaiter(_this, void 0, void 0, function () {
                var sheetName, sheets, sheet, entireSheet, countOfNumbers, numbers, i, range;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            logger.clear();
                            logger.comment("Creating test data ...");
                            sheetName = "PerfData";
                            sheets = ctx.workbook.worksheets;
                            sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
                            ctx.load(sheet);
                            return [4 /*yield*/, ctx.sync()];
                        case 1:
                            _a.sent();
                            if (!sheet.isNull) return [3 /*break*/, 3];
                            sheet = sheets.add(sheetName);
                            ctx.load(sheet);
                            return [4 /*yield*/, ctx.sync()];
                        case 2:
                            _a.sent();
                            _a.label = 3;
                        case 3:
                            entireSheet = sheet.getRange(null);
                            entireSheet.clear(null);
                            return [4 /*yield*/, ctx.sync];
                        case 4:
                            _a.sent();
                            countOfNumbers = 10000;
                            numbers = [];
                            for (i = countOfNumbers; i > 0; i--) {
                                numbers.push([i]);
                            }
                            range = sheet.getRange('A1:A' + countOfNumbers);
                            range.values = numbers;
                            return [4 /*yield*/, ctx.sync()];
                        case 5:
                            _a.sent();
                            logger.comment("Test data has been created.");
                            return [2 /*return*/];
                    }
                });
            }); });
            return [2 /*return*/];
        });
    });
}
var logger = {};
window.onload = function () {
    var loggerElement = document.getElementById('log');
    logger.comment = function () {
        for (var i = 0; i < arguments.length; i++) {
            if (typeof arguments[i] == 'object') {
                loggerElement.innerHTML += (JSON && JSON.stringify ? JSON.stringify(arguments[i], undefined, 2) : arguments[i]) + '<br />';
            }
            else {
                loggerElement.innerHTML += arguments[i] + '<br />';
            }
        }
    };
    logger.clear = function () {
        loggerElement.innerHTML = '';
    };
};
function runTest() {
    var _this = this;
    Excel.run(function (ctx) { return __awaiter(_this, void 0, void 0, function () {
        var numberOfIterations, numberOfFunctions, numberOfFunctionsOnEachRow, timeout, selectedFunction, radios, i, radioButton, functionName, functionParameters, bubbleSortNumbers, nthPrime, principalAmount, interestRate, numberOfMonths, err_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    logger.clear();
                    numberOfIterations = parseInt(document.getElementById('NumberOfIterations').value);
                    numberOfFunctions = parseInt(document.getElementById('NumberOfFunctions').value);
                    numberOfFunctionsOnEachRow = parseInt(document.getElementById('NumberOfFunctionsOnEachRow').value);
                    timeout = parseInt(document.getElementById('Timeout').value);
                    selectedFunction = "0";
                    radios = document.getElementsByName('FunctionToRun');
                    for (i = 0; i < radios.length; i++) {
                        radioButton = radios[i];
                        if (radioButton.checked) {
                            selectedFunction = radioButton.value;
                            break;
                        }
                    }
                    functionName = '';
                    functionParameters = '';
                    switch (selectedFunction) {
                        case "1":
                            functionName = functionNamespace + "bubbleSortJS";
                            bubbleSortNumbers = document.getElementById('BubbleSortNumbers').value;
                            functionParameters = "(PerfData!A1:A" + bubbleSortNumbers + ")";
                            break;
                        case "2":
                            functionName = functionNamespace + "findNthPrimeJS";
                            nthPrime = document.getElementById('NthPrime').value;
                            functionParameters = "(" + nthPrime + ")";
                            break;
                        default:
                            functionName = functionNamespace + "mortgagePaymentJS";
                            principalAmount = document.getElementById('PrincipalAmount').value;
                            interestRate = document.getElementById('InterestRate').value;
                            numberOfMonths = document.getElementById('NumberOfMonths').value;
                            functionParameters = "(" + principalAmount + "," + interestRate + "," + numberOfMonths + ")";
                            break;
                    }
                    logger.comment("Calling =" + functionName + functionParameters);
                    return [4 /*yield*/, customFunctions_AsyncPerfHelper(ctx, numberOfIterations, numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters, timeout, 1 /*expectedEventCount*/)];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    err_2 = _a.sent();
                    logger.comment("Failure occurred: " + err_2);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    }); });
    return 0;
}
;
function promisify(action) {
    return new Promise(function (resolve, reject) {
        var callback = function (result) {
            if (result.status === "succeeded") {
                resolve(result.value);
            }
            else {
                reject(result.error);
            }
        };
        action(callback);
    });
}
function customFunctions_Async_Execute(ctx, arrayOfFormulas, expectedTimeout, expectedEventCount) {
    return __awaiter(this, void 0, void 0, function () {
        var testApi, endTimestamp, beginTimestamp, beginEventFired, endEventFired, numberOfRows, numberOfCols, eventCounter, entireSheet;
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    testApi = ctx.workbook.internalTest;
                    beginEventFired = false;
                    endEventFired = false;
                    numberOfRows = arrayOfFormulas.length;
                    numberOfCols = arrayOfFormulas[0].length;
                    eventCounter = 0;
                    // unregister all custom function events
                    testApi.unregisterAllCustomFunctionExecutionEvents();
                    entireSheet = ctx.workbook.worksheets.getItem('Sheet1').getRange(null);
                    entireSheet.clear(null);
                    return [4 /*yield*/, ctx.sync()];
                case 1:
                    _a.sent();
                    return [2 /*return*/, promisify(function (callback) { return __awaiter(_this, void 0, void 0, function () {
                            var beginEventListener, timeoutHandle, endEventListener, columnName, range;
                            var _this = this;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        beginEventListener = testApi.onCustomFunctionExecutionBeginEvent.add(function (eventArgs) {
                                            if (!beginEventFired) {
                                                beginTimestamp = customFunctions_ConvertTicksToMicroseconds(eventArgs.higherTicks, eventArgs.lowerTicks);
                                                beginEventFired = true;
                                            }
                                            eventCounter++;
                                            if (eventCounter == expectedEventCount) {
                                                beginEventListener.remove();
                                                return ctx.sync();
                                            }
                                        });
                                        endEventListener = testApi.onCustomFunctionExecutionEndEvent.add(function (eventArgs) {
                                            if (eventCounter == expectedEventCount) {
                                                clearTimeout(timeoutHandle);
                                                endTimestamp = customFunctions_ConvertTicksToMicroseconds(eventArgs.higherTicks, eventArgs.lowerTicks);
                                                endEventListener.remove();
                                                endEventFired = true;
                                                var executionTimeMicroseconds = 0;
                                                if (beginEventFired) {
                                                    executionTimeMicroseconds = endTimestamp - beginTimestamp;
                                                }
                                                callback({ status: 'succeeded', value: executionTimeMicroseconds });
                                                return ctx.sync();
                                            }
                                        });
                                        return [4 /*yield*/, ctx.sync()];
                                    case 1:
                                        _a.sent();
                                        columnName = getColumnName(numberOfCols);
                                        range = ctx.workbook.worksheets.getItem('Sheet1').getRange('A1:' + columnName + numberOfRows);
                                        range.formulas = arrayOfFormulas;
                                        timeoutHandle = setTimeout(function () { return __awaiter(_this, void 0, void 0, function () {
                                            return __generator(this, function (_a) {
                                                switch (_a.label) {
                                                    case 0:
                                                        if (!beginEventFired) {
                                                            beginEventListener.remove();
                                                        }
                                                        endEventListener.remove();
                                                        return [4 /*yield*/, ctx.sync()];
                                                    case 1:
                                                        _a.sent();
                                                        callback({ status: 'failed', error: "Customfunction execution event(s) were not fired, CustomFunctionExecutionBeginEvent=" + beginEventFired + ", CustomFunctionExecutionEndEvent=" + endEventFired });
                                                        return [2 /*return*/];
                                                }
                                            });
                                        }); }, expectedTimeout);
                                        return [4 /*yield*/, ctx.sync()];
                                    case 2:
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        }); })];
            }
        });
    });
}
function createAddRangeInputData(ctx, addRangeInputParameter, isIntergerData) {
    return __awaiter(this, void 0, void 0, function () {
        var inputParameterSplitIndex, range, addRangeInput, rangeSplit, columnMax, rowMax, rangeInputData, i, j;
        return __generator(this, function (_a) {
            inputParameterSplitIndex = addRangeInputParameter.indexOf("!");
            range = addRangeInputParameter.substr(inputParameterSplitIndex + 1);
            addRangeInput = ctx.workbook.worksheets.getItem(addRangeInputParameter.substr(0, inputParameterSplitIndex)).getRange(range);
            if (isIntergerData) {
                rangeSplit = range.indexOf(":");
                columnMax = 20;
                rowMax = parseInt(range.substr(rangeSplit + 2));
                rangeInputData = new Array(rowMax);
                for (i = 0; i < rowMax; i++) {
                    rangeInputData[i] = new Array(columnMax);
                    for (j = 0; j < columnMax; j++) {
                        rangeInputData[i][j] = 167 * Math.sin((i + j) * 97);
                    }
                }
                addRangeInput.values = rangeInputData;
            }
            else {
                addRangeInput.values = "abcdefghijklmnopqrst";
            }
            return [2 /*return*/, ctx.sync()];
        });
    });
}
function customFunctions_Sync_Execute(ctx, arrayOfFormulas) {
    var testApi = ctx.workbook.internalTest;
    var endTimestamp;
    var beginTimestamp;
    var beginEventFired = false;
    var endEventFired = false;
    var numberOfRows = arrayOfFormulas.length;
    var numberOfCols = arrayOfFormulas[0].length;
    var eventCounter = 0;
    // clear the worksheet
    var entireSheet = ctx.workbook.worksheets.getItem('Sheet1').getRange(null);
    entireSheet.clear(null);
    ctx.sync();
    // execute the functions
    var columnName = getColumnName(numberOfCols);
    var range = ctx.workbook.worksheets.getItem('Sheet1').getRange('A1:' + columnName + numberOfRows);
    range.formulas = arrayOfFormulas;
    ctx.sync();
}
function customFunctions_AsyncPerfHelper(ctx, numberOfIterations, numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters, expectedTimeout, expectedEventCount) {
    return __awaiter(this, void 0, void 0, function () {
        var arrayOfFormulas, executionTimes, i, executionTime, avgExecutionTime, medianExecutionTime;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    arrayOfFormulas = createFormulas(numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters);
                    executionTimes = [];
                    i = 0;
                    _a.label = 1;
                case 1:
                    if (!(i < numberOfIterations)) return [3 /*break*/, 4];
                    return [4 /*yield*/, customFunctions_Async_Execute(ctx, arrayOfFormulas, expectedTimeout, expectedEventCount)];
                case 2:
                    executionTime = (_a.sent()) / 1000;
                    logger.comment("Iteration " + (i + 1) + ": Total execution time " + executionTime + " ms");
                    avgExecutionTime = executionTime / numberOfFunctions;
                    logger.comment("Iteration " + (i + 1) + ": Avg. time per function " + avgExecutionTime + " ms");
                    executionTimes.push(avgExecutionTime);
                    _a.label = 3;
                case 3:
                    i++;
                    return [3 /*break*/, 1];
                case 4:
                    executionTimes.sort(function (a, b) { return a - b; });
                    medianExecutionTime = 0;
                    if (numberOfIterations % 2 == 0) {
                        medianExecutionTime = (executionTimes[(numberOfIterations / 2) - 1] + executionTimes[numberOfIterations / 2]) / 2;
                    }
                    else {
                        medianExecutionTime = executionTimes[(numberOfIterations - 1) / 2];
                    }
                    logger.comment("Median (avg. time per function) = " + medianExecutionTime + " ms");
                    return [2 /*return*/, medianExecutionTime];
            }
        });
    });
}
function customFunctions_SyncPerfHelper(ctx, numberOfIterations, numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters) {
    var arrayOfFormulas = createFormulas(numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters);
    var executionTimes = [];
    for (var i = 0; i < numberOfIterations; i++) {
        customFunctions_Sync_Execute(ctx, arrayOfFormulas);
    }
}
function createFormulas(numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters) {
    var arrayOfFormulas = [];
    var formula = '=' + functionName + functionParameters;
    do {
        var formulasForARow = [];
        for (var i = 0; i < numberOfFunctionsOnEachRow; i++) {
            if (numberOfFunctions > 0) {
                numberOfFunctions--;
                formulasForARow.push(formula);
            }
            else {
                formulasForARow.push("");
            }
        }
        arrayOfFormulas.push(formulasForARow);
    } while (numberOfFunctions > 0);
    return arrayOfFormulas;
}
function customFunctions_ConvertTicksToMicroseconds(higherTicks, lowerTicks) {
    // We cannot use << here because, the result of << is always a 32bit integer
    var microseconds = Math.pow(2, 31) * higherTicks + lowerTicks;
    return microseconds;
}
function getColumnName(columnNumber) {
    var columnName = "";
    var temp = 0;
    while (columnNumber > 0) {
        temp = columnNumber % 26;
        if (temp == 0) {
            temp = 26;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        else {
            columnNumber = Math.floor(columnNumber / 26);
        }
        columnName = String.fromCharCode(temp + 64) + columnName;
    }
    return columnName;
}
