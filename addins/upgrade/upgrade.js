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

function getSharedValue() {
	if (typeof(g_sharedState) === 'object') {
		return g_sharedState.value;
	}

	return null;
}

function setSharedValue(value) {
	if (typeof(g_sharedState) === 'object') {
		g_sharedState.value = value;
		return value;
	}

	return null;
}

function getRuntimeState() {
	return fficeRuntime.CurrentRuntime.getState();
}

function setRuntimeState(value) {
	return OfficeRuntime.CurrentRuntime.setState(value);
}

CustomFunctions.associate('VERSIONSYNC', getVersion);
CustomFunctions.associate('VERSIONASYNC', getVersion);
CustomFunctions.associate('VERSIONDELAYED', function(ms) { return delay(getVersion, ms); });

CustomFunctions.associate('CONSTSYNC', getConst);
CustomFunctions.associate('CONSTASYNC', getConst);
CustomFunctions.associate('CONSTDELAYED', function(ms) { return delay(getConst, ms); });

CustomFunctions.associate('STREAMSEQUENCE', streamSequence);

CustomFunctions.associate('GETSHAREDVALUE', getSharedValue);
CustomFunctions.associate('SETSHAREDVALUE', setSharedValue);

CustomFunctions.associate('GETRUNTIMESTATE', getRuntimeState);
CustomFunctions.associate('SETRUNTIMESTATE', setRuntimeState);
