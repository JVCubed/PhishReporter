'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            getPhishingItem(Office.context.mailbox.item);
            composeMail();
        });
    });

    // defining global variables to pass them to the composeMail function
    var phishItemId;
    var phishSubject

    // this function has to run first because when the new message form pops up the current iteminformation is lost. 
    function getPhishingItem(item) {
        phishItemId = item.itemId
        phishSubject = item.subject
    }

    function composeMail() {
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: ["Jeroen.van.der.Velden2@hva.nl"],
            // ccRecipients: ["sam@contoso.com"], Send to more mailaddresses if nessecery
            subject: Office.context.mailbox.userProfile.displayName + " is reporting phishing!",
            htmlBody: 'Hello <b>phishing</b>!<br/></i>' + phishItemId,
            attachments: [
                { type: "item", itemId: phishItemId, name: phishSubject }
            ],
        });
    }

})();