var appContext = {};
window.addEventListener("message", receiveMessage, false);
var XdmMessagePackager = (function () {
    function XdmMessagePackager() {
    }
    XdmMessagePackager.envelope = function (messageObject, serializerVersion) {
        if (typeof (messageObject) === "object") {
            messageObject._serializerVersion = 1;
        }
        return JSON.stringify(messageObject);
    };
    XdmMessagePackager.unenvelope = function (messageObject, serializerVersion) {
        return JSON.parse(messageObject);
    };
    return XdmMessagePackager;
}());
function receiveMessage(e) {
    if (e.data != '') {
        var messageObject;
        var serializerVersion = 1;
        var serializedMessage = e.data;
        try {
            messageObject = XdmMessagePackager.unenvelope(serializedMessage, 1);
            serializerVersion = messageObject._serializerVersion != null ? messageObject._serializerVersion : serializerVersion;
        }
        catch (ex) {
            return;
        }
        if (messageObject._actionName == "HostAppContextAsync") {
            appContext = messageObject;
            appContext._actionName = "ContextActivationManager_getAppContextAsync";
        }
        else if (messageObject._actionName == "ContextActivationManager_getAppContextAsync")
        {
            var requestJson = XdmMessagePackager.unenvelope(e.data, 1);            
            appContext._conversationId = requestJson._conversationId;
            e.source.postMessage(XdmMessagePackager.envelope(appContext, 1), e.origin);
        }
    }
}