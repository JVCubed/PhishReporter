'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
       

        $(document).ready(function () {
            // The document is ready
            loadCurrentMailAddress();  
        });
    });

})();

function loadCurrentMailAddress() {
    // Write message property values to the task pane
    document.getElementById("currentMailAddress").innerHTML = Office.context.roamingSettings.get("email");
}

function changeMailAddress() {
    var newMailAddress = document.getElementById("newMailAddress").value;
    Office.context.roamingSettings.set("email", newMailAddress);
    saveRoamingSettings();
}

function saveRoamingSettings() {
    // Save settings in the mailbox to make it available in future sessions.
    Office.context.roamingSettings.saveAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error(`Action failed with message ${result.error.message}`);
        } else {
            console.log(`Settings saved with status: ${result.status}`);
            loadCurrentMailAddress()
        }
    });
}