/// <reference path="../Office.Runtime.js" />
/// <reference path="../Excel.js" />

function getVersion() {
	return "18.7.24.2";
}

function getConst() {
	return 42;
}

function streamSequence(init, step, count, ms, context) {
	if (count > 0) {
		delay(function() {
			context.setResult(init);
			streamSequence(init + step, step, count - 1, ms, context);
		}, ms);
	}
}

function delay(func, ms) {
	return new Promise(function(resolve) {
		setTimeout(function() {
			resolve(func());
		}, ms);
	});
}

CustomFunctionMappings = {
	VERSION_SYNC: getVersion,
	VERSION_ASYNC: getVersion,
	VERSION_DELAYED: function(ms) { return delay(getVersion, ms); },

	CONST_SYNC: getConst,
	CONST_ASYNC: getConst,
	CONST_DELAYED: function(ms) { return delay(getConst, ms); },

	STREAM_SEQUENCE: streamSequence
};
