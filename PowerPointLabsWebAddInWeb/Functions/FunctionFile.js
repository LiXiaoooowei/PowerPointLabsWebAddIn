// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

function animateInSlide(event) {
    Office.context.document.setSelectedDataAsync('Animate In Slide Button tapped',
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showNotification("Error", asyncResult.error.message);
            }
        });
    event.completed();
}

function addAnimationSlide(event) {
    Office.context.document.setSelectedDataAsync('Add Animation Slide Button tapped',
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showNotification("Error", asyncResult.error.message);
            }
        });
    event.completed();
}

function openSettingsDialog(event) {
    Office.context.ui.displayDialogAsync(window.location.origin + "/AnimationSettingsDialog.html", { height: 50, width: 50, displayInIframe: true },
        function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Failed) {
                var dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                    dialog.close();
                    console.log(arg.message);
                });
            }
        });
    event.completed();
}
