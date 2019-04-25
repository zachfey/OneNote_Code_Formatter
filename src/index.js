import * as OfficeHelpers from "@microsoft/office-js-helpers";

Office.onReady(() => {
    // Office is ready
    $(document).ready(() => {
        // The document is ready
        $('#addOutline').click(addOutlineToPage);
    });
});

async function addOutlineToPage() {
    try {
        await OneNote.run(async context => {
            var html = "<p>" + $("#textBox").val() + "</p>";

            // Get the current page.
            var page = context.application.getActivePage();

            // Queue a command to load the page with the title property.
            page.load("title");

            // Add text to the page by using the specified HTML.
            var outline = page.addOutline(40, 90, html);

            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function() {
                    console.log("Added outline to page " + page.title);
                })
                .catch(function(error) {
                    app.showNotification("Error: " + error);
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
            });
    } catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    }
}