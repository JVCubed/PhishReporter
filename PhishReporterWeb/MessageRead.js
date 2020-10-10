'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
            
            composeMail();
        });
    });

    

    function loadItemProps(item) {
        // Write message property values to the task pane
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-internetMessageId').text(item.internetMessageId);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");

        

        

        //$('#debug').text(item.itemClass);

       

        //if (item.displayNewMessageForm !== undefined) {
        //    $('#debug').text("displayNewMessageForm can be used.");
            // Use item.displayNewMessageForm.
        //    item.displayNewMessageForm;
        //} else {
        //    $('#debug').text("displayNewMessageForm can't be used.");
        //}
    }

    function composeMail() {
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: Office.context.mailbox.item.to, // Copies the To line from current item
            ccRecipients: ["sam@contoso.com"],
            subject: "Outlook add-ins are cool!",
            htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
        });
    }

})();