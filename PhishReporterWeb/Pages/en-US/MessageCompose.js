﻿'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            securityTeamMailAddress();
            loadCurrentMailAddress()
            getPhishingItem(Office.context.mailbox.item);
            composeMail();
        });
    });

    // defining global variables to pass them to the composeMail function
    var phishItemId;
    var phishSubject;
    var receipentMailAddress;


    // get the reciepent or ask to enter value
    // check if there is an email address set to send the mail to
    function securityTeamMailAddress() {
        // check if email is already set
        if (Office.context.roamingSettings.get("email")) {
            receipentMailAddress = Office.context.roamingSettings.get("email")
        }
        // show popup to enter the email address to report phishing to.
        else {
            // TODO: Create popup to enter email at first run

            // Office.context.roamingSettings.set("email", "j.vdvelden99@gmail.com")
            receipentMailAddress = Office.context.roamingSettings.get("email")
            saveRoamingSettings()
        }   
    }

    // save value's to roaming settings so it can be accessed later
    function saveRoamingSettings() {
        // Save settings in the mailbox to make it available in future sessions.
        Office.context.roamingSettings.saveAsync(function (result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.error(`Action failed with message ${result.error.message}`);
            } else {
                console.log(`Settings saved with status: ${result.status}`);
            }
        });
    }


    // this function has to run before composing a new mail to retrieve the details of the current selected email. 
    function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    // function to open a new 'compose message' form with predefined information
    function composeMail() {
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: [receipentMailAddress],
            // ccRecipients: ["sam@contoso.com"], Send to more mailaddresses if necessary
            subject: "Phishing report: \"" + phishSubject + "\"",
            htmlBody:
                'Dear Support,<br/><br/>' +
                'I received attached email and want to report it as phishing.' +
                '<br/><br/>Please write down any additional information below to line.' +
                ' e.g., that you\'ve clicked on a link (hopefully not).' +
                '<br/>--------------------------------------------------',
            attachments: [
                { type: "item", itemId: phishItemId, name: phishSubject }
            ],
        });
    }

})();

function hideShowSettings() {
    if (document.getElementById("settings").style.display === "none") {
        document.getElementById("settings").style.display = "block";
    } else {
        document.getElementById("settings").style.display = "none";
    };
};

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