namespace WordAddIn {
    interface StringMap {
        [index: string]: string;
    }
    interface ContentControl {
        tag: string;
        title: string;
    }

    class DemoController {

        constructor() {
            this.getAllContentControls();
        }

        getContentControls(): void {
            // Run a batch operation against the Word object model.
            Word.run(function (context) {
	
                // Create a proxy object for the content controls collection.
                var contentControls = context.document.contentControls;
	
                // Queue a command to load the id property for all of content controls. 
                context.load(contentControls, 'id');
	 
                // Synchronize the document state by executing the queued-up commands, 
                // and return a promise to indicate task completion.
                return context.sync().then((): void => {
                    // Show the id property for each content control.
                    for (var i = 0; i < contentControls.items.length; i++) {
                        console.log("contentControl[" + i + "].id = " +
                            contentControls.items[i].id);
                    }
                });
            })
                .catch((error: any): void => {
                    console.log('Error: ' + JSON.stringify(error));
                    if (error instanceof OfficeExtension.Error) {
                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                    }
                });
        }

        // Helper that deduplicates the set of content controls. Content controls are
        // considered duplicates if they share the same tag.
        removeDuplicateContentControls(contentControls: Word.ContentControlCollection): ContentControl[] {

            let i: number;
            let len: number = contentControls.items.length;
            let uniqueFields: ContentControl[] = [];
            let currentContentControl: StringMap = {};

            for (i = 0; i < len; i++) {
                currentContentControl[contentControls.items[i].tag] = contentControls.items[i].title;
            }

            let tag: string;
            for (tag in currentContentControl) {

                let obj: ContentControl = {
                    tag: tag,
                    title: currentContentControl[tag]
                };

                uniqueFields.push(obj);
            }

            return uniqueFields;
        }
       


        // Using the Word JS API. Gets all of the content controls that are in the loaded document. 
        getAllContentControls(): void {
            Word.run(function (context: Word.RequestContext) {

                // Create a proxy object for the document.
                let thisDocument: Word.Document = context.document;

                // Create a proxy object for the content control collection in the current document.
                let contentControls: Word.ContentControlCollection = thisDocument.contentControls;

                // Queue a command to load the tag properties of the content controls.
                contentControls.load("tag");

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync<Word.ContentControlCollection>(contentControls).then((): void => {

                });
            })
                .catch((error: any) => {
                    console.log("Error: " + JSON.stringify(error));
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
        }
    }

    angular.module("app")
        .controller("app.DemoController", DemoController);
}