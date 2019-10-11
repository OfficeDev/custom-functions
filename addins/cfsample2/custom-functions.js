function promiseWorkaround(func) {
	var timerId;
	if (typeof(document) !== 'undefined' && document.createElement){
		timerId = setInterval(function(){
			var div = document.createElement('div');
			div.innerHTML = 'Promise Workaround';
			div.style.display = 'none';
			document.body.appendChild(div);
		}, 20);
	}

	return func()
		.then(function(value) {
			if (timerId) {
				clearInterval(timerId);
			}

			return value;
		});
}

function SetRuntimeStateHelper(state) {
	return promiseWorkaround(function () {
		return OfficeRuntime.currentRuntime.setState(state)
			.then(function () {
				if (typeof(g_BgAppRuntimeState) === 'object') {
					g_BgAppRuntimeState.value = state;
				}
				return state;
			})
			.catch(function (error) {
				return error.code;
			});
	});
}

function SetStartupStateHelper(state) {
	return promiseWorkaround(function () {
		return OfficeRuntime.currentRuntime.setStartupState(state)
			.then(function () {
				if (typeof(g_BgAppRuntimeStartupState) === 'object') {
					g_BgAppRuntimeStartupState.value = state;
				}
				return state;
			})
			.catch(function (error) {
				return error.code;
			});
	});
}

CustomFunctions.associate('GetCFDataRangeValue', function(address){
	return promiseWorkaround(function() {
		var context = new Excel.RequestContext();
		var sheet = context.workbook.worksheets.getItemOrNullObject("CFData");
		var range;
		return context.sync()
			.then(function() {
				if (sheet.isNullObject) {
					sheet = context.workbook.worksheets.add("CFData");
				}
				return context.sync();
			})
			.then(function() {
				range = sheet.getRange(address);
				range.load();
				return context.sync();
			})
			.then(function() {
				return range.values[0][0];
			});
	});
});

CustomFunctions.associate('GetRangeValue', function(address){
	return promiseWorkaround(function() {
		var context = new Excel.RequestContext();
		var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
		range.load();
		return context.sync()
			.then(function() {
				return range.values[0][0];
			});
	});
});


CustomFunctions.associate('SetRangeValue', function(address, value){
	return promiseWorkaround(function() {
		var context = new Excel.RequestContext();
		var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
		range.values = [[value]];
		range.load();
		return context.sync()
			.then(function() {
				return range.values[0][0];
			});
	});
});

CustomFunctions.associate('GetValue', function() {
	if (typeof(g_sharedAppData) === 'object') {
		return g_sharedAppData.value;
	}

	return null;
});
CustomFunctions.associate('SetValue', function(value) {
	if (typeof(g_sharedAppData) === 'object') {
		g_sharedAppData.value = value;
		return value;
	}

	return null;
});

CustomFunctions.associate('GetRuntimeState', function() {
	return promiseWorkaround(function () {
		return OfficeRuntime.currentRuntime.getState()
		.then(function (value) {
			if (typeof(g_BgAppRuntimeState) === 'object') {
				g_BgAppRuntimeState.value = value;
			}
			return value.toString();
		})
		.catch(function (error) {
			return error.code;
		});
	});
});

CustomFunctions.associate('SetRuntimeState', function (state) {
	return SetRuntimeStateHelper(state);
});

CustomFunctions.associate('GetStartupState', function() {
	return promiseWorkaround(function () {
		return OfficeRuntime.currentRuntime.getStartupState()
		.then(function (value) {
			if (typeof(g_BgAppRuntimeStartupState) === 'object') {
				g_BgAppRuntimeStartupState.value = value;
			}
			return value.toString();
		})
		.catch(function (error) {
			return error.code;
		});
	});
});



CustomFunctions.associate('SetStartupState', function (state) {
	return SetStartupStateHelper(state);
});
