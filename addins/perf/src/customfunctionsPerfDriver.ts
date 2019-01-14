let functionNamespace: string = "msft.perf.";

async function customFunctions_runAutomatedTests() {
	Excel.run(async (ctx) => {

		let inputSheetName: string = "Async Input";
		let inputDataRangeName: string = "InputTests";
		let range: Excel.Range = ctx.workbook.worksheets.getItem(inputSheetName).getRange(inputDataRangeName);
		ctx.load(range);
		await ctx.sync();

		let outputSheet: string;
		let radios = document.getElementsByName('OutputSheet');
		for (let i = 0; i < radios.length; i++) {
			let radioButton: HTMLInputElement = <HTMLInputElement>radios[i];
			if (radioButton.checked) {
				outputSheet = radioButton.value;
				break;
			}
		}

		let numberOfFunctions: number = 0;
		let expectedTimeout: number = 0;
		let numberOfIterations: number = 0;
		let functionName: string;
		let functionId: number = 0;
		let functionParameters: string;
		let numberOfFunctionsOnEachRow: number = 0;
		let testName: string;
		let expectedEventCount: number = 0;
		let outputData: any[][] = [];
		outputData.push(["Test Id", "Test Name", "Avg. time (ms) per function"]);

		logger.clear();

		// we will skip the first row, as that is for header
		for (let i = 1; i < range.rowCount; i++) {
			// This order must be kept in sync with the columns in the 'Input' sheet in the perf xlsx file
			functionId = range.values[i][0];
			functionName = functionNamespace + range.values[i][1];
			numberOfFunctions = range.values[i][2];
			numberOfIterations = range.values[i][3];
			numberOfFunctionsOnEachRow = range.values[i][4];
			expectedTimeout = range.values[i][5];
			functionParameters = '(' + range.values[i][6] + ')';
			expectedEventCount = range.values[i][7] + 1;

			if (range.values[i][1] == "addRangeJS" || range.values[i][1] == "syncAddRangeJS")
			{
				let addRangeInputParameter: string = range.values[i][6];
				await createAddRangeInputData(ctx, addRangeInputParameter, true);
			}

			if (range.values[i][1] == "addStringRangeJS" || range.values[i][1] == "syncAddStringRangeJS") {
				let addRangeInputParameter: string = range.values[i][6];
				await createAddRangeInputData(ctx, addRangeInputParameter, false);
			}

			testName = range.values[i][1] + '_' + numberOfFunctions;
			if (expectedEventCount > 1) {
				testName += '_' + 'chainedCalls' + '_' + (expectedEventCount - 1);
			}
			try {
				logger.comment("Running test [id = " + functionId + ", function = " + functionName + functionParameters + " ]");
				let executionTime = await customFunctions_runAutomatedTestByTestIdHelper(ctx, range, i, true);
				outputData.push([functionId, testName, executionTime]);
			}
			catch (err) {
				logger.comment("Error:" + JSON.stringify(err));
				outputData.push([functionId, testName, "Error:" + JSON.stringify(err)]);
			}
		}

		logger.comment("Writting results");
		// clear the entire sheet
		let entireSheet = ctx.workbook.worksheets.getItem(outputSheet).getRange(null);
		entireSheet.clear(null);

		// write the output to the specifed output sheet
		let outputRange: Excel.Range = ctx.workbook.worksheets.getItem(outputSheet).getRange('A1:C' + range.rowCount);
		outputRange.values = outputData;
		await ctx.sync();
	})
	.catch((ex) => {
		logger.comment("Error: " + JSON.stringify(ex));
	});
}

function customFunctions_runAutomatedTestByTestId(testId: number) {
	Excel.run(async (ctx) => {
		let isAsyncExecutionElement : HTMLInputElement = document.getElementById("IsAsyncExecution") as HTMLInputElement;
		let isAsyncExecution: boolean = isAsyncExecutionElement.checked;
		let isStressTestElement: HTMLInputElement = document.getElementById("IfExecuteStressTest") as HTMLInputElement;
		let isStressTest: boolean = isStressTestElement.checked;
		let inputSheetName: string = isStressTest ? (isAsyncExecution ? "Async Stress Test Input" : "Sync Stress Test Input" ) : (isAsyncExecution ? "Async Input" : "Sync Input");
		let inputDataRangeName: string = "InputTests";
		let range: Excel.Range = ctx.workbook.worksheets.getItem(inputSheetName).getRange(inputDataRangeName);
		ctx.load(range);
		await ctx.sync();

		customFunctions_runAutomatedTestByTestIdHelper(ctx, range, testId, isAsyncExecution);
	})
	.catch((ex) => {
		logger.comment("Error: " + JSON.stringify(ex));
	});
}

async function customFunctions_runAutomatedTestByTestIdHelper(ctx: any, range: Excel.Range, testId: number, isAsyncExecution: boolean)
{
	let numberOfFunctions: number = 0;
	let expectedTimeout: number = 0;
	let numberOfIterations: number = 0;
	let functionName: string;
	let functionId: number = 0;
	let functionParameters: string;
	let numberOfFunctionsOnEachRow: number = 0;
	let expectedEventCount: number = 0;
	let outputData: any[][] = [];

	let executionTime: number = 0;

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

	if (range.values[testId][1] == "addRangeJS" || range.values[testId][1] == "syncAddRangeJS") {
		let addRangeInputParameter: string = range.values[testId][6];
		await createAddRangeInputData(ctx, addRangeInputParameter, true);
	}

	if (range.values[testId][1] == "addStringRangeJS" || range.values[testId][1] == "syncAddStringRangeJS") {
		let addRangeInputParameter: string = range.values[testId][6];
		await createAddRangeInputData(ctx, addRangeInputParameter, false);
	}
	if (isAsyncExecution) {
		executionTime = await customFunctions_AsyncPerfHelper(ctx, numberOfIterations, numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters, expectedTimeout, expectedEventCount);
	}
	else
	{
		customFunctions_SyncPerfHelper(ctx, numberOfIterations, numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters);
	}

	return executionTime;
}

async function customFunctions_createTestData() {
	Excel.run(async (ctx) => {
		logger.clear();
		logger.comment("Creating test data ...");
		let sheetName: string = "PerfData";

		let sheets = ctx.workbook.worksheets;
		let sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
		ctx.load(sheet);
		await ctx.sync();

		if (sheet.isNull) {
			sheet = sheets.add(sheetName);
			ctx.load(sheet);
			await ctx.sync();
		}

		let entireSheet = sheet.getRange(null);
		entireSheet.clear(null);
		await ctx.sync;

		const countOfNumbers: number = 10000;
		let numbers = [];
		for (let i = countOfNumbers; i > 0 ; i--){
			numbers.push([i]);
		}

		let range: Excel.Range = sheet.getRange('A1:A'+ countOfNumbers);
		range.values = numbers;
		await ctx.sync();
		logger.comment("Test data has been created.");
	});
}

let logger:any = {};

window.onload = function() { 
	let loggerElement: HTMLElement = document.getElementById('log');

	logger.comment = function () {
		for (var i = 0; i < arguments.length; i++) {
		if (typeof arguments[i] == 'object') {
			loggerElement.innerHTML += (JSON && JSON.stringify ? JSON.stringify(arguments[i], undefined, 2) : arguments[i]) + '<br />';
		} else {
			loggerElement.innerHTML += arguments[i] + '<br />';
		}
		}
	}

	logger.clear = function () {
		loggerElement.innerHTML = '';
	}
};

function runTest(): number {
	Excel.run(async (ctx) => {
		try
		{
			logger.clear();
			let numberOfIterations: number = parseInt((<HTMLInputElement>document.getElementById('NumberOfIterations')).value);
			let numberOfFunctions: number = parseInt((<HTMLInputElement>document.getElementById('NumberOfFunctions')).value);
			let numberOfFunctionsOnEachRow: number = parseInt((<HTMLInputElement>document.getElementById('NumberOfFunctionsOnEachRow')).value);
			let timeout: number = parseInt((<HTMLInputElement>document.getElementById('Timeout')).value);
			let selectedFunction: string = "0";
			let radios = document.getElementsByName('FunctionToRun');
			for(let i = 0; i < radios.length; i++) {
				let radioButton: HTMLInputElement = <HTMLInputElement>radios[i];
				if(radioButton.checked) {
					selectedFunction = radioButton.value;
					break;
				}
			}

			let functionName: string = '';
			let functionParameters: string = '';
			switch(selectedFunction) {
				case "1":
					functionName = functionNamespace + "bubbleSortJS";
					let bubbleSortNumbers:string = (<HTMLInputElement>document.getElementById('BubbleSortNumbers')).value;
					functionParameters = "(PerfData!A1:A" + bubbleSortNumbers + ")";
					break;
				case "2":
					functionName = functionNamespace + "findNthPrimeJS";
					let nthPrime:string = (<HTMLInputElement>document.getElementById('NthPrime')).value;
					functionParameters = "(" + nthPrime + ")";
					break;

				default:
					functionName = functionNamespace + "mortgagePaymentJS";
					let principalAmount: string = (<HTMLInputElement>document.getElementById('PrincipalAmount')).value;
					let interestRate: string = (<HTMLInputElement>document.getElementById('InterestRate')).value;
					let numberOfMonths: string = (<HTMLInputElement>document.getElementById('NumberOfMonths')).value;
					functionParameters = "(" + principalAmount + "," + interestRate + "," + numberOfMonths + ")";
					break;
			}
			logger.comment("Calling ="  + functionName + functionParameters);

			await customFunctions_AsyncPerfHelper(ctx, numberOfIterations, numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters, timeout, 1 /*expectedEventCount*/);
		}
		catch (err)
		{
			logger.comment("Failure occurred: " + err);
		}
	});

	return 0;
};

function promisify<T>(action: (callback) => void): Promise<T> {
	return new Promise(function (resolve, reject) {
		var callback = function (result) {
			if (result.status === "succeeded") {
				resolve(result.value);
			} else {
				reject(result.error);
			}
		};

		action(callback);
	});
}

async function customFunctions_Async_Execute(ctx: any, arrayOfFormulas: number[][], expectedTimeout: number, expectedEventCount: number): Promise<number> {
	let testApi: Excel.InternalTest = ctx.workbook.internalTest;
	let endTimestamp: number;
	let beginTimestamp: number;
	let beginEventFired: boolean = false;
	let endEventFired: boolean = false;
	let numberOfRows: number = arrayOfFormulas.length;
	let numberOfCols: number = arrayOfFormulas[0].length;
	let eventCounter = 0;

	// unregister all custom function events
	testApi.unregisterAllCustomFunctionExecutionEvents();	
	// clear the worksheet
	let entireSheet = ctx.workbook.worksheets.getItem('Sheet1').getRange(null);
	entireSheet.clear(null);
	await ctx.sync();

	return promisify<number>(async (callback) => {
		let beginEventListener: OfficeExtension.EventHandlerResult<Excel.CustomFunctionEventArgs> = testApi.onCustomFunctionExecutionBeginEvent.add((eventArgs: Excel.CustomFunctionEventArgs) => {
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

		let timeoutHandle: number;
		let endEventListener: OfficeExtension.EventHandlerResult<Excel.CustomFunctionEventArgs> = testApi.onCustomFunctionExecutionEndEvent.add((eventArgs: Excel.CustomFunctionEventArgs) => {
			if (eventCounter == expectedEventCount) {
				clearTimeout(timeoutHandle);
				endTimestamp = customFunctions_ConvertTicksToMicroseconds(eventArgs.higherTicks, eventArgs.lowerTicks);
				endEventListener.remove();
				endEventFired = true;

				let executionTimeMicroseconds: number = 0;
				if (beginEventFired) {
					executionTimeMicroseconds =  endTimestamp - beginTimestamp;
				}
				callback( {status: 'succeeded', value: executionTimeMicroseconds} );
				return ctx.sync();
			}
		});
		await ctx.sync();

		// execute the  functions
		let columnName: string = getColumnName(numberOfCols);
		let range: Excel.Range = ctx.workbook.worksheets.getItem('Sheet1').getRange('A1:' + columnName + numberOfRows);
		range.formulas = arrayOfFormulas;

		timeoutHandle = setTimeout(async () => {
			if (!beginEventFired) {
				beginEventListener.remove();
			}
			endEventListener.remove();
			await ctx.sync();

			callback( {status: 'failed', error: "Customfunction execution event(s) were not fired, CustomFunctionExecutionBeginEvent=" + beginEventFired +", CustomFunctionExecutionEndEvent=" + endEventFired} );
		}, expectedTimeout);

		await ctx.sync();
	});
}

async function createAddRangeInputData(ctx: any, addRangeInputParameter: string, isIntergerData: boolean)
{
	let inputParameterSplitIndex = addRangeInputParameter.indexOf("!");
	let range = addRangeInputParameter.substr(inputParameterSplitIndex + 1); 
	let addRangeInput: Excel.Range = ctx.workbook.worksheets.getItem(addRangeInputParameter.substr(0, inputParameterSplitIndex)).getRange(range);
	if (isIntergerData) {
		// The input range gaven in the excel test workbook is A1:Tx.
		// Therefore, the column number 20 predefined.
		// The row number need to get from x
		let rangeSplit = range.indexOf(":");
		let columnMax = 20;
		let rowMax = parseInt(range.substr(rangeSplit + 2)); // Row number is after ":T"
		let rangeInputData = new Array(rowMax);
		for (let i: number = 0; i < rowMax; i++)
		{
			rangeInputData[i] = new Array(columnMax);
			for (let j: number  = 0; j < columnMax; j++)
			{
				rangeInputData[i][j] = 167 * Math.sin((i + j) * 97);
			}
		}
		addRangeInput.values = rangeInputData;
	}
	else
	{
		addRangeInput.values = <any> "abcdefghijklmnopqrst";
	}
	return ctx.sync();
}

function customFunctions_Sync_Execute(ctx: any, arrayOfFormulas: number[][]) {
	let testApi: Excel.InternalTest = ctx.workbook.internalTest;
	let endTimestamp: number;
	let beginTimestamp: number;
	let beginEventFired: boolean = false;
	let endEventFired: boolean = false;
	let numberOfRows: number = arrayOfFormulas.length;
	let numberOfCols: number = arrayOfFormulas[0].length;
	let eventCounter = 0;

	// clear the worksheet
	let entireSheet = ctx.workbook.worksheets.getItem('Sheet1').getRange(null);
	entireSheet.clear(null);
	ctx.sync();

	// execute the functions
	let columnName: string = getColumnName(numberOfCols);
	let range: Excel.Range = ctx.workbook.worksheets.getItem('Sheet1').getRange('A1:' + columnName + numberOfRows);
	range.formulas = arrayOfFormulas;

	ctx.sync();
}

async function customFunctions_AsyncPerfHelper(ctx: any, numberOfIterations: number, numberOfFunctions: number, numberOfFunctionsOnEachRow: number, functionName: string, functionParameters: string, expectedTimeout: number, expectedEventCount: number) {
	let arrayOfFormulas: number[][] = createFormulas(numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters);
	let executionTimes = [];
	for (var i =0; i < numberOfIterations; i++) {
		let executionTime = await customFunctions_Async_Execute(ctx, arrayOfFormulas, expectedTimeout, expectedEventCount)/1000;
		logger.comment("Iteration " + (i+1) + ": Total execution time " + executionTime + " ms");
		let avgExecutionTime = executionTime/numberOfFunctions;
		logger.comment("Iteration " + (i+1) + ": Avg. time per function " + avgExecutionTime + " ms");
		executionTimes.push(avgExecutionTime);
	}
	executionTimes.sort((a, b) => a - b);

	let medianExecutionTime: number = 0;
	if (numberOfIterations % 2 == 0){
		medianExecutionTime = (executionTimes[(numberOfIterations/2) - 1] + executionTimes[numberOfIterations/2])/2;
	} else{
		medianExecutionTime = executionTimes[(numberOfIterations-1)/2];
	}

	logger.comment("Median (avg. time per function) = " + medianExecutionTime + " ms");
	return medianExecutionTime;
}

function customFunctions_SyncPerfHelper(ctx: any, numberOfIterations: number, numberOfFunctions: number, numberOfFunctionsOnEachRow: number, functionName: string, functionParameters: string) {
	let arrayOfFormulas: number[][] = createFormulas(numberOfFunctions, numberOfFunctionsOnEachRow, functionName, functionParameters);
	let executionTimes = [];
	for (var i = 0; i < numberOfIterations; i++) {
		customFunctions_Sync_Execute(ctx, arrayOfFormulas);
	}
}

function createFormulas(numberOfFunctions: number, numberOfFunctionsOnEachRow: number, functionName: string, functionParameters: string): number[][] {
	let arrayOfFormulas = [];
	let formula: string = '=' + functionName + functionParameters;
	do {
		let formulasForARow = [];
		for (let i: number = 0; i < numberOfFunctionsOnEachRow; i++){
			if (numberOfFunctions > 0) {
				numberOfFunctions--;
				formulasForARow.push(formula);
			} else {
				formulasForARow.push("");
			}
		}
		arrayOfFormulas.push(formulasForARow);
	} while (numberOfFunctions > 0);

	return arrayOfFormulas;
}

function customFunctions_ConvertTicksToMicroseconds(higherTicks: number, lowerTicks: number) : number {
	// We cannot use << here because, the result of << is always a 32bit integer
	let microseconds:number = Math.pow(2,31) * higherTicks + lowerTicks;
	return microseconds;
}

function getColumnName(columnNumber: number): string {
	let columnName: string = "";
	let temp: number = 0;
	while (columnNumber > 0) {
		temp = columnNumber % 26;
		if (temp == 0) {
			temp = 26;
			columnNumber = Math.floor((columnNumber - 1) / 26);
		} else {
			columnNumber = Math.floor(columnNumber / 26);
		}
		columnName = String.fromCharCode(temp+64) + columnName;
	}

	return columnName;
}