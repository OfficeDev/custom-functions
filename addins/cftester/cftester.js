(function () {
	"use strict";

	var cellToHighlight;
	var messageBanner;

	// The initialize function must be run each time a new page is loaded.
	Office.initialize = function (reason) {
		$(document).ready(function () {
			// Initialize the FabricUI notification mechanism and hide it
			var element = document.querySelector('.ms-MessageBanner');
			messageBanner = new fabric.MessageBanner(element);
			messageBanner.hideBanner();

			// Add a click event handler for the highlight button.
			$('#highlight-button').click(hightlightHighestValue);
			$("#recalc").click(recalc);
			$("#get-result").click(getResult);
			$("#run").click(runFunction);
		});
	};

	function recalc() {
		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {
			// Create a proxy object for the selected range and load its properties
			ctx.workbook.application.calculate(Excel.CalculationType.full);
			return ctx.sync();
		})
		.catch(errorHandler);
	}

	function getResult() {
		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {
			let address = document.getElementById("target-cell").value;
			let range = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);
			range.load("values");
			// Run the queued-up command, and return a promise to indicate task completion
			return ctx.sync()
				.then(function () {
					document.getElementById("result").textContent = range.values[0];
					showNotification('The range.values is:', '"' + range.values + '"');
				})
				.then(ctx.sync);
		})
			.catch(errorHandler);
	}

	function runFunction() {
		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {
			let address = document.getElementById("target-cell").value;
			let arrayOfFormulas = document.getElementById("function").value;
			let range = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);
			range.formulas = arrayOfFormulas;
			range.load("formulas");
			return ctx.sync()
				.then(function () {
					console.log(range.formulas);
					showNotification('The range.formulas is:', '"' + range.formulas + '"');
				})
				.then(ctx.sync);
		})
		.catch(errorHandler);
	}

	// Helper function for treating errors
	function errorHandler(error) {
		// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
		showNotification("Error", error);
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
	}

	// Helper function for displaying notifications
	function showNotification(header, content) {
		$("#notification-header").text(header);
		$("#notification-body").text(content);
		messageBanner.showBanner();
		messageBanner.toggleExpansion();
	}
})();
