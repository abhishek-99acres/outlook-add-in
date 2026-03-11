// FIX 4: Office.actions.associate is REQUIRED to map the manifest function
// name to the actual JS handler. Without this the event never fires.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

let globalEvent = null;
let dialogRef = null;

function onMessageSendHandler(event) {

    const item = Office.context.mailbox.item;

    // No attachments → allow send immediately
    if (item.attachments.length === 0) {
        event.completed({ allowEvent: true });
        return;
    }

    item.loadCustomPropertiesAsync(function(result) {

        const props = result.value;
        const categorized = props.get("attachmentsCategorized");

        if (!categorized) {

            globalEvent = event;

            Office.context.ui.displayDialogAsync(
                "https://abhishek-99acres.github.io/outlook-add-in/dialog.html",
                { height: 40, width: 40, promptBeforeOpen: false },
                function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        event.completed({
                            allowEvent: false,
                            errorMessage: "Could not open categorization dialog. Please try again."
                        });
                        return;
                    }

                    dialogRef = asyncResult.value;

                    dialogRef.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                        dialogRef.close();
                        if (arg.message === "categorized") {
                            globalEvent.completed({ allowEvent: true });
                        } else {
                            globalEvent.completed({
                                allowEvent: false,
                                errorMessage: "Please categorize attachments before sending."
                            });
                        }
                    });

                    dialogRef.addEventHandler(Office.EventType.DialogEventReceived, function() {
                        globalEvent.completed({
                            allowEvent: false,
                            errorMessage: "Please categorize attachments before sending."
                        });
                    });
                }
            );

        } else {
            event.completed({ allowEvent: true });
        }
    });
}
