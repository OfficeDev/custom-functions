Office.actions.associate('SHOWTASKPANE', function () {
	return Office.addin.showAsTaskpane();
});

Office.actions.associate('HIDETASKPANE', function () {
	return Office.addin.hide();
});

Office.actions.associate('SETBOLD', function () {
	var context = new Excel.RequestContext();
	var range = context.workbook.getSelectedRange();
	var rangeFormat = range.format;
	rangeFormat.font.load();
	return context.sync()
		.then(function () {
			var rangeTarget = context.workbook.getSelectedRange();
			if (rangeFormat.font.bold) {
				rangeTarget.format.font.bold = false;
			} else {
				rangeTarget.format.font.bold = true;
			}
			return context.sync();
		});
});

Office.actions.associate('SETITALIC', function () {
	var context = new Excel.RequestContext();
	var range = context.workbook.getSelectedRange();
	var rangeFormat = range.format;
	rangeFormat.font.load();
	return context.sync()
		.then(function () {
			var rangeTarget = context.workbook.getSelectedRange();
			if (rangeFormat.font.italic) {
				rangeTarget.format.font.italic = false;
			} else {
				rangeTarget.format.font.italic = true;
			}
			return context.sync();
		});
});

Office.actions.associate('SETUNDERLINE', function () {
	var context = new Excel.RequestContext();
	var range = context.workbook.getSelectedRange();
	var rangeFormat = range.format;
	rangeFormat.font.load();
	return context.sync()
		.then(function () {
			var rangeTarget = context.workbook.getSelectedRange();
			if (rangeFormat.font.underline !== "None") {
				rangeTarget.format.font.underline = "None";
			} else {
				rangeTarget.format.font.underline = "Single";
			}
			return context.sync();
		});
});

Office.actions.associate('SETCOLOR', function () {
	var context = new Excel.RequestContext();
	var range = context.workbook.getSelectedRange();
	var rangeFormat = range.format;
	rangeFormat.fill.load();
	var colors = ["#FFFFFF", "#C7CC7A", "#7560BA", "#9DD9D2", "#FFE1A8", "#E26D5C"];
	return context.sync().then(function () {
		var rangeTarget = context.workbook.getSelectedRange();
		var currentColor = -1;
		for (var i = 0; i < colors.length; i++) {
			if (colors[i] == rangeFormat.fill.color) {
				currentColor = i;
				break;
			}
		}
		if (currentColor == -1) {
			currentColor = 0;
		}
		else if (currentColor == colors.length - 1) {
			currentColor = 0;
		}
		else {
			currentColor++;
		}
		rangeTarget.format.fill.color = colors[currentColor];
		return context.sync();
	});
});

Office.actions.associate('SETDATEFORMAT', function () {
	var context = new Excel.RequestContext();
	var range = context.workbook.getSelectedRange();
	range.load();
	return context.sync().then(function () {
		var rangeFormat = range.numberFormat;
		var dateFormats = ["m/d/yy", "mm/dd/yy", "mm/dd/yyyy", "m/d/yyyy", "ddd mm/dd/yyyy", "ddddd, mmm dd, yyyy"];
		var result = [];
		var rangeTarget = context.workbook.getSelectedRange();
		var currentFormat = -1;
		for (var i = 0; i < dateFormats.length; i++) {
			if (dateFormats[i] == rangeFormat[0][0]) {
				currentFormat = i;
				break;
			}
		}
		if (currentFormat == -1) {
			currentFormat = 0;
		} else if (currentFormat == dateFormats.length - 1) {
			currentFormat = 0;
		} else {
			currentFormat++;
		}
		for (var j = 0; j < range.rowCount; j++) {
			result[j] = [];
			for (var n = 0; n < range.columnCount; n++) {
				result[j][n] = dateFormats[currentFormat];
			}
		}
		rangeTarget.numberFormat = result;
		return context.sync();
	});
});
