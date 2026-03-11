Office.onReady(() => {});

// Global reference to hold the dialog and event between async calls
let globalEvent = null;
let dialogRef = null;

function onMessageSendHandler(event) {

    const item = Office.context.mailbox.item;

    // No attachments → allow send immediately
    if (item.attachments.length === 0) {
        event.completed({ allowEvent: true });
        return;
    }

    // Check if we need to filter by To/CC recipients
    // Remove the comment block below if you want to restrict to specific addresses:
    /*
    const targetAddress = "specific@example.com";
    const allRecipients = [
        ...item.to.map(r => r.emailAddress),
        ...item.cc.map(r => r.emailAddress)
    ];
    const isTargetted = allRecipients.some(addr =>
        addr.toLowerCase() === targetAddress.toLowerCase()
    );
    if (!isTargetted) {
        event.completed({ allowEvent: true });
        return;
    }
    */

    item.loadCustomPropertiesAsync(function(result) {

        const props = result.value;
        const categorized = props.get("attachmentsCategorized");

        if (!categorized) {

            // Save event reference — we'll complete it after dialog responds
            globalEvent = event;

            Office.context.ui.displayDialogAsync(
                "https://YOUR-HOSTED-DOMAIN/dialog.html",
                { height: 40, width: 40, promptBeforeOpen: false },
                function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        // If dialog fails to open, block the send with a message
                        event.completed({
                            allowEvent: false,
                            errorMessage: "Could not open categorization dialog. Please try again."
                        });
                        return;
                    }

                    dialogRef = asyncResult.value;

                    // Listen for messages sent back from dialog.js via messageParent()
                    dialogRef.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                        dialogRef.close();

                        if (arg.message === "categorized") {
                            // User completed categorization — allow the send
                            globalEvent.completed({ allowEvent: true });
                        } else {
                            // User cancelled or something went wrong — block send
                            globalEvent.completed({
                                allowEvent: false,
                                errorMessage: "Please categorize attachments before sending."
                            });
                        }
                    });

                    dialogRef.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
                        // User closed the dialog manually (e.g. clicked X)
                        globalEvent.completed({
                            allowEvent: false,
                            errorMessage: "Please categorize attachments before sending."
                        });
                    });
                }
            );

            // DO NOT call event.completed() here — we wait for dialog to respond above

        } else {
            // Already categorized in a previous attempt — allow send
            event.completed({ allowEvent: true });
        }
    });
}
