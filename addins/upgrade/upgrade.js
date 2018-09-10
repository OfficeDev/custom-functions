/// <reference path="../Office.Runtime.js" />
/// <reference path="../Excel.js" />

function version() {
	return "18.7.24.2";
}

function sleep(ms) {
	return new Promise(function(resolve) {
		setTimeout(function() {
			resolve(version());
		}, ms);
	});
}

CustomFunctionMappings = {
	VERSION_SYNC: version,
	VERSION_ASYNC: version,
	SLEEP: sleep
};
