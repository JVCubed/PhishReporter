'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            securityTeamMailAddress();
            getPhishingItem(Office.context.mailbox.item);
            composeMail();
        });
    });

    // defining global variables to pass them to the composeMail function
    var phishItemId;
    var phishSubject;
    var securityTeamMailAddress;

    // check if there is an email address set to send the mail to
    function securityTeamMailAddress() {
        // check if email is already set
        if (Office.context.roamingSettings.get("email")) {
            securityTeamMailAddress = Office.context.roamingSettings.get("email")
        }
        // show popup to enter the email address to report phishing to.
        else {
            // TODO: Create popup to enter email at first run

            // Office.context.roamingSettings.set("email", "j.vdvelden99@gmail.com")
            securityTeamMailAddress = Office.context.roamingSettings.get("email")
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
            toRecipients: [securityTeamMailAddress],
            // ccRecipients: ["sam@contoso.com"], Send to more mailaddresses if nessecery
            subject: Office.context.mailbox.userProfile.displayName + " is reporting phishing!",
            htmlBody: 'PhishItemId: ' + phishItemId + '<br>MailAddress: ' + securityTeamMailAddress,
            attachments: [
                { type: "item", itemId: phishItemId, name: phishSubject }
            ],
        });
    }

})();