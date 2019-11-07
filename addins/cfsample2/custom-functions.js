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

function SetRuntimeVisibleHelper(visible) {
	return promiseWorkaround(function () {
		var p;
		if (visible) {
			p = OfficeRuntime.currentTaskpane.show();
		}
		else {
			p = OfficeRuntime.currentTaskpane.hide();
		}

		return p.then(function () {
				return visible;
			})
			.catch(function (error) {
				return error.code;
			});
	});
}

function SetStartupBehaviorHelper(state) {
	return promiseWorkaround(function () {
		return OfficeRuntime.currentTaskpane.setStartupBehavior(state)
			.then(function () {
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
		// _getState() is the internal API and it's only for Microsoft engineer team internal testing purpose. Please do not use it.
		return OfficeRuntime.currentTaskpane._getState().then(function (value) {
			return value;
		})
		.catch(function (error) {
			return error.code;
		});
	});
});

CustomFunctions.associate('GetVisibility', function() {
	return g_BgAppVisibilityState.value;
});

CustomFunctions.associate('Show', function () {
	return SetRuntimeVisibleHelper(true);
});

CustomFunctions.associate('Hide', function () {
	return SetRuntimeVisibleHelper(false);
});

CustomFunctions.associate('GetStartupBehavior', function() {
	return promiseWorkaround(function () {
		return OfficeRuntime.currentTaskpane.getStartupBehavior()
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



CustomFunctions.associate('SetStartupBehavior', function (behavior) {
	return SetStartupBehaviorHelper(behavior);
});
