var WordAddIn;
(function (WordAddIn) {
    var DemoController = (function () {
        function DemoController() {
            this.getAllContentControls();
        }
        DemoController.prototype.getContentControls = function () {
            // Run a batch operation against the Word object model.
            Word.run(function (context) {
                // Create a proxy object for the content controls collection.
                var contentControls = context.document.contentControls;
                // Queue a command to load the id property for all of content controls. 
                context.load(contentControls, 'id');
                // Synchronize the document state by executing the queued-up commands, 
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    // Show the id property for each content control.
                    for (var i = 0; i < contentControls.items.length; i++) {
                        console.log("contentControl[" + i + "].id = " +
                            contentControls.items[i].id);
                    }
                });
            })
                .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        };
        // Helper that deduplicates the set of content controls. Content controls are
        // considered duplicates if they share the same tag.
        DemoController.prototype.removeDuplicateContentControls = function (contentControls) {
            var i;
            var len = contentControls.items.length;
            var uniqueFields = [];
            var currentContentControl = {};
            for (i = 0; i < len; i++) {
                currentContentControl[contentControls.items[i].tag] = contentControls.items[i].title;
            }
            var tag;
            for (tag in currentContentControl) {
                var obj = {
                    tag: tag,
                    title: currentContentControl[tag]
                };
                uniqueFields.push(obj);
            }
            return uniqueFields;
        };
        // Using the Word JS API. Gets all of the content controls that are in the loaded document. 
        DemoController.prototype.getAllContentControls = function () {
            Word.run(function (context) {
                // Create a proxy object for the document.
                var thisDocument = context.document;
                // Create a proxy object for the content control collection in the current document.
                var contentControls = thisDocument.contentControls;
                // Queue a command to load the tag properties of the content controls.
                contentControls.load("tag");
                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync(contentControls).then(function () {
                });
            })
                .catch(function (error) {
                console.log("Error: " + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        };
        return DemoController;
    })();
    angular.module("app")
        .controller("app.DemoController", DemoController);
})(WordAddIn || (WordAddIn = {}));
//# sourceMappingURL=demo.controller.js.map