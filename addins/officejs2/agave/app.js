_perfData.appJsExecutionStart = performance.now();

var _bodyOnLoadCalled = false;
var _pendingLogs = [];
function BodyOnLoad() {
    _bodyOnLoadCalled = true;
    for (var i = 0; i < _pendingLogs.length; i++) {
        log(_pendingLogs[i]);
    }
}

function log(text) {
    if (!_bodyOnLoadCalled) {
        _pendingLogs.push(text);
        return;
    }
    var div = document.createElement('div');
    div.appendChild(document.createTextNode(text));
    document.getElementById('DivLog').appendChild(div);
}

Office.onReady(function (hostAndPlatform) {
    _perfData.officeOnReadyApp = performance.now();
    _perfData.officeOnReadyAppDuration = _perfData.officeOnReadyApp - _perfData.start;
    log('_perfData');
    log(JSON.stringify(_perfData));
    
    if (typeof(OSFPerformance) !== 'undefined') {
        log('OSFPerformance');
        log(JSON.stringify(OSFPerformance));
        var summary = {
            networkDuration: OSFPerformance.officeExecuteStart - _perfData.start,
            officeJsExecutionDuration: OSFPerformance.officeExecuteEnd - OSFPerformance.officeExecuteStart,
            officeJsStartToGetAppContext: OSFPerformance.getAppContextStart - OSFPerformance.officeExecuteStart,
            getAppContextDuration: OSFPerformance.getAppContextEnd - OSFPerformance.getAppContextStart,
            getAppContextXdmDuration: OSFPerformance.getAppContextXdmEnd - OSFPerformance.getAppContextXdmStart,
            officeOnReadyDuration: Math.max(OSFPerformance.officeExecuteEnd, OSFPerformance.officeOnReady) - OSFPerformance.officeExecuteStart,
            officeOnReadyAppDuration: _perfData.officeOnReadyAppDuration
        };
        log('Summary:');
        log(JSON.stringify(summary));
    }

    log(Office.context.displayLanguage);
    log(hostAndPlatform.host);
    log(hostAndPlatform.platform);
    var isSupported = Office.context.requirements.isSetSupported('ExcelApi', '1.7');
    log('ExcelApi1.7=' + isSupported);
});

Office.onReady(function (hostAndPlatform) {
    log('[SecondOnReady] platform=' + hostAndPlatform.platform);
});

function BtnRunClick() {
    var code = document.getElementById("TxtCode").value;
    eval(code);
}

function BtnReloadClick() {
    window.location.reload(true);
}
function BtnTestClick() {
    var context = new Excel.RequestContext();
    var sheet = context.workbook.worksheets.getItem('Sheet1');
    var r = sheet.getRange('A1');
    var text = 'Hello' + Math.random();
    r.values = [[text]];
    context.sync()
        .then(function () {
            log('Done');
            log(sheet.id);
        })
        .catch(function (ex) {
            log(JSON.stringify(ex));
        });
}
function BtnTest2Click() {
    var context = new Excel.RequestContext();
    var sheet = context.workbook.worksheets.getItem('Sheet1');
    var r = sheet.getRange('A1:C20');
    var text = 'Hello' + Math.random();
    r.values = text;
    r.load();
    context.sync()
        .then(function () {
            log('Done');
            log(sheet.id);
            log(JSON.stringify(r));
        })
        .catch(function (ex) {
            log(JSON.stringify(ex));
        });
}

function BtnTestEventClick() {
    Excel.run(function (ctx) {
        var testApi = ctx.workbook.internalTest;
        var eventResult = testApi.onTestEvent.add(function (eventArgs) {
            log("Test Event triggered");
            log(JSON.stringify(eventArgs));
            return null;
        });

        var eventResult1 = testApi.onTest1Event.add(function (eventArgs) {
            log("Test Event1 triggered");
            log(JSON.stringify(eventArgs));
            return null;
        });
        return ctx.sync()
            .then(function () {
                log("Trigger events");
                testApi.triggerTestEventWithFilter(200, 1 /*Excel.MessageType.testEvent*/, ctx.workbook.worksheets.getFirst());
                testApi.triggerTestEventWithFilter(201, 2 /*Excel.MessageType.test1Event*/, ctx.workbook.worksheets.getFirst());
                return ctx.sync();
            })
            .then(function () {
                log("Waiting for 10 seconds...");
                return OfficeExtension.Utility._createTimeoutPromise(10000);
            })
            .then(function () {
                log("After 10 seconds, calling workbook.load()...");
                ctx.workbook.load();
                return ctx.sync();
            })
            .then(function () {
                log("After 10 seconds and sync, removing event handlers...");
                eventResult.remove()
                eventResult1.remove();
                return ctx.sync();
            })
            .then(function () {
                log("Done");
            })
            .catch(function (ex) {
                log("Error:" + JSON.stringify(ex));
            });
    });
}

function test_settings_updateUsingV2() {
    Excel.run(function (ctx) {
        ctx.workbook.settings.add('stringKey', 'Hello');
        // ctx.workbook.settings.add('intKey', 1000);
        // ctx.workbook.settings.add('dateKey', new Date());
        ctx.sync()
            .then(function () {
                log("Done");
            })
            .catch(function (ex) {
                log("Error:" + JSON.stringify(ex));
            });
    });
}

function test_settings_readV1() {
    var stringValue = Office.context.document.settings.get('stringKey');
    var intValue = Office.context.document.settings.get('intKey');
    var dateValue = Office.context.document.settings.get('dateKey');
    log(stringValue);
    log(intValue);
    log(dateValue);
}
function BtnClearLogClick() {
    document.getElementById('DivLog').innerHTML = '';
}


function appcmdTestButton(args) {
    log('appcmdTestButton invoked');
    args.completed();
}

_perfData.appJsExecutionEnd = performance.now();
