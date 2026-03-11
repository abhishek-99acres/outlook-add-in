function saveAndSend() {

    const category = document.getElementById("category").value;
    const item = Office.context.mailbox.item;

    item.loadCustomPropertiesAsync(function(result) {

        const props = result.value;

        props.set("attachmentsCategorized", true);
        props.set("attachmentCategoryValue", category);

        props.saveAsync(function() {
            // FIXED: item.sendAsync() does not exist.
            // Instead, message back to the parent (events.js) which holds the event reference.
            // The parent will then call event.completed({ allowEvent: true }) to proceed with send.
            Office.context.ui.messageParent("categorized");
        });
    });
}

function cancel() {
    Office.context.ui.messageParent("cancelled");
}
