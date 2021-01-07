function getSubject() {
    Office.context.mailbox.item.subject.getAsync(callback);

    function callback(asyncResult) {
        var subject = asyncResult.value;
        writeResult("GetSubject result is: "+ subject);
    }
}

function setSubject() {
    Office.context.mailbox.item.subject.setAsync("New subject!", function (asyncResult) {
        if (asyncResult.status === "failed") {
            console.log("Action failed with error: " + asyncResult.error.message);
            writeResult("Set subject Action failed with error: " + asyncResult.error.message);
        }
        else {
            writeResult("SetSubject done.");
        }
    });
}

function getUserProfile() {
    writeResult(Office.context.mailbox.userProfile.displayname);
}

function getRestId() {
    let itemId = Office.context.mailbox.item.itemId;
    let restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    writeResult(restId);
}

function getLanguage() {
    writeResult(Office.context.displayLanguage);
}

function closeContainer() {
    Office.context.ui.closeContainer();
}

function clearResult() {
    document.getElementById("result").value = "";
}

function writeResult(result) {
    document.getElementById("result").value += ("\n" +result);
}

Office.initialize = function (reason) {};
Office.onReady(function(info) {});