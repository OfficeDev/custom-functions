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
    log(JSON.stringify(_perfData));
    if (typeof(OSFPerformance) !== 'undefined') {
        log('OSFPerformance');
        log(JSON.stringify(OSFPerformance));
    }

    log(Office.context.displayLanguage);
    log(hostAndPlatform.host);
    log(hostAndPlatform.platform);
    // var isSupported = Office.context.requirements.isSetSupported('ExcelApi', '1.7');
    // log('ExcelApi1.7=' + isSupported);
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
    Word.run(function (context) {
    
        var textSample = 'This is an example of the insert text method. This is a method ' + 
            'which allows users to insert text into a selection. It can insert text into a ' +
            'relative location or it can overwrite the current selection. Since the ' +
            'getSelection method returns a range object, look up the range object documentation ' +
            'for everything you can do with a selection.';
        
        // Create a range proxy object for the current selection.
        var range = context.document.getSelection();
        
        // Queue a command to insert text at the end of the selection.
        range.insertText(textSample, Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            log('Inserted the text at the end of the selection.');
        });  
    })
    .catch(function (error) {
        log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
    
}

function BtnClearLogClick() {
    document.getElementById('DivLog').innerHTML = '';
}