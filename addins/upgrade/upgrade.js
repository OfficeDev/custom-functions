/// <reference path="../Office.Runtime.js" />
/// <reference path="../Excel.js" />

function getVersion() {
	return "18.7.24.2";
}

function getConst() {
	return 42;
}

function _delay(func, ms) {
	return new Promise(function(resolve) {
		setTimeout(function() {
			resolve(func());
		}, ms);
	});
}

CustomFunctionMappings = {
	VERSION_SYNC: getVersion,
	VERSION_ASYNC: getVersion,
	VERSION_DELAYED: function(ms) { return _delay(getVersion, ms); },

	CONST_SYNC: getConst,
	CONST_ASYNC: getConst,
	CONST_DELAYED: function(ms) { return _delay(getConst, ms); }
};
