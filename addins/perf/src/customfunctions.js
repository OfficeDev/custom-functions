function swap(numbers, i, j) {
    var temp = numbers[i];
    numbers[i] = numbers[j];
    numbers[j] = temp;
}
function mortgagePaymentJS(principalAmount, interestRate, numberOfMonths) {
    interestRate = (interestRate / 100) / 12;
    var irr = Math.pow(1 + interestRate, numberOfMonths);
    return interestRate * principalAmount * irr / (irr - 1);
}
function bubbleSortJS(array) {
    var numbers = [];
    for (var i = 0; i < array.length; ++i) {
        for (var j = 0; j < array[i].length; ++j) {
            numbers.push(array[i][j]);
        }
    }
    var numberOfSwaps = 0;
    var length = numbers.length;
    for (var i = 0; i < length - 1; i++) {
        for (var j = 0; j < length - i - 1; j++) {
            if (numbers[j] > numbers[j + 1]) {
                numberOfSwaps++;
                swap(numbers, j, j + 1);
            }
        }
    }
    return numberOfSwaps;
}
function isPrimeNumber(current, vectorPrimes) {
    var length = vectorPrimes.length;
    for (var i = 0; i < length; i++) {
        if ((current % vectorPrimes[i]) == 0) {
            return false;
        }
    }
    return true;
}
function findNthPrimeJS(n) {
    if (n <= 0) {
        return 0;
    }
    var count = 0;
    var current = 1;
    var vectorPrimes = [];
    do {
        current++;
        if (isPrimeNumber(current, vectorPrimes)) {
            vectorPrimes.push(current);
            count++;
        }
    } while (count < n);
    return current;
}
function addRangeJS(range) {
    var sum = 0;
    for (var i = 0; i < range.length; ++i) {
        for (var j = 0; j < range[i].length; ++j) {
            sum += range[i][j];
        }
    }
    return sum;
}
function addStringRangeJS(range) {
    var sum = 0;
    for (var i = 0; i < range.length; ++i) {
        for (var j = 0; j < range[i].length; ++j) {
            sum += range[i][j].length;
        }
    }
    return sum;
}
function nestedLoopJS(n) {
    var sum = 0;
    for (var i = 0; i < n; ++i) {
        for (var j = 0; j < n; ++j) {
            sum++;
        }
    }
    return sum;
}
CustomFunctionMappings.MORTGAGEPAYMENTJS = mortgagePaymentJS;
CustomFunctionMappings.SYNCMORTGAGEPAYMENTJS = mortgagePaymentJS;
CustomFunctionMappings.FINDNTHPRIMEJS = findNthPrimeJS;
CustomFunctionMappings.SYNCFINDNTHPRIMEJS = findNthPrimeJS;
CustomFunctionMappings.BUBBLESORTJS = bubbleSortJS;
CustomFunctionMappings.SYNCBUBBLESORTJS = bubbleSortJS;
CustomFunctionMappings.ADDRANGEJS = addRangeJS;
CustomFunctionMappings.SYNCADDRANGEJS = addRangeJS;
CustomFunctionMappings.ADDSTRINGRANGEJS = addStringRangeJS;
CustomFunctionMappings.SYNCADDSTRINGRANGEJS = addStringRangeJS;
CustomFunctionMappings.NESTEDLOOPJS = nestedLoopJS;
CustomFunctionMappings.SYNCNESTEDLOOPJS = nestedLoopJS;
