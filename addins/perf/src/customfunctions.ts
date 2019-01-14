function swap(numbers: number[], i: number, j: number): void {
	let temp: number = numbers[i];
	numbers[i] = numbers[j];
	numbers[j] = temp;
}

function mortgagePaymentJS(principalAmount: number, interestRate: number, numberOfMonths: number): number {
	interestRate = (interestRate / 100) / 12;
	let irr: number = Math.pow(1 + interestRate, numberOfMonths);
	return interestRate * principalAmount * irr / (irr - 1);
}

function bubbleSortJS(array: Array<Array<number>>) : number {
	let numbers: number[] = [];
	for (let i = 0; i < array.length; ++i) {
		for (let j = 0; j < array[i].length; ++j) {
			numbers.push(array[i][j]);
		}
	}

	let numberOfSwaps: number = 0;
	let length: number = numbers.length;
	for (let i: number = 0; i < length - 1; i++)	{
		for (let j: number = 0; j < length - i - 1; j++) {
			if (numbers[j] > numbers[j + 1]) {
				numberOfSwaps++;
				swap(numbers, j, j + 1);
			}
		}
	}
	return numberOfSwaps;
}

function isPrimeNumber(current: number, vectorPrimes: number[]): boolean {		
	let length: number = vectorPrimes.length;
	for (let i: number = 0; i < length; i++) {
		if ((current % vectorPrimes[i]) == 0) {
			return false;
		}
	}
	return true;
}

function findNthPrimeJS(n: number): number {
	if (n <= 0) {
		return 0;
	}

	let count: number = 0;
	let current: number = 1;
	let vectorPrimes: number[] = [];
	do {
		current++;
		if (isPrimeNumber(current, vectorPrimes)) {
			vectorPrimes.push(current);
			count++;
		}
	} while (count < n);

	return current;
}

function addRangeJS(range: number[][]): number{
	var sum = 0;
	for (var i = 0; i < range.length; ++i) {
		for (var j = 0; j < range[i].length; ++j) {
			sum += range[i][j];
		}
	}

	return sum;
}

function addStringRangeJS(range: string[][]): number {
	var sum = 0;
	for (var i = 0; i < range.length; ++i) {
		for (var j = 0; j < range[i].length; ++j) {
			sum += range[i][j].length;
		}
	}

	return sum;
}

function nestedLoopJS(n: number): number {
	var sum = 0;
	for (var i = 0; i < n; ++i) {
		for (var j = 0; j < n; ++j) {
			sum ++;
		}
	}

	return sum;
}

CustomFunctions.associate("mortgagePaymentJS", mortgagePaymentJS);
CustomFunctions.associate("syncMortgagePaymentJS", mortgagePaymentJS);
CustomFunctions.associate("findNthPrimeJS", findNthPrimeJS);
CustomFunctions.associate("syncFindNthPrimeJS", findNthPrimeJS);
CustomFunctions.associate("bubbleSortJS", bubbleSortJS);
CustomFunctions.associate("syncBubbleSortJS", bubbleSortJS);
CustomFunctions.associate("addRangeJS", addRangeJS);
CustomFunctions.associate("syncAddRangeJS", addRangeJS);
CustomFunctions.associate("addStringRangeJS", addStringRangeJS);
CustomFunctions.associate("syncAddStringRangeJS", addStringRangeJS);
CustomFunctions.associate("nestedLoopJS", nestedLoopJS);
CustomFunctions.associate("syncNestedLoopJS", nestedLoopJS);
