// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $("#okButton").click(handleSubmit);
            $("#cancelButton").click(handleCancel);
        });
    };

    function handleSubmit() {
        Office.context.ui.messageParent("submit");
    }
    function handleCancel() {
        Office.context.ui.messageParent("cancel");
    }
})();