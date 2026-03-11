function saveAndSend() {

    const category = document.getElementById("category").value;
    const item = Office.context.mailbox.item;

    item.loadCustomPropertiesAsync(function(result) {

        const props = result.value;

        props.set("attachmentsCategorized", true);
        props.set("attachmentCategoryValue", category);

        props.saveAsync(function() {
            Office.context.ui.messageParent("categorized");
        });
    });
}

function cancel() {
    Office.context.ui.messageParent("cancelled");
}
