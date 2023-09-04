'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
            $("#btnAPI").unbind("click");
            $("#btnAPI").click(function () {
                //let  = 'https://assistance.azurewebsites.net/api/Priorite/GetAll';
                //const response = await fetch(url);
                //const myJson = await response.json(); //extract JSON from the http response
                console.log("BtnAPI");
            });
        });
    });

    function loadItemProps(item) {
        console.log("Test2");
        // Write message property values to the task pane
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-internetMessageId').text(item.internetMessageId);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
    }
});