/// <reference path="../../typings/tsd.d.ts" />

namespace WordAddIn {

    interface INotification {
        header: string;
        body: string;
    }


    interface IHomeController {
        header: string;
        showNotification: boolean;
        hideNotification(): void;
        getDataFromSelection(): void;
    }

   
    class HomeController implements IHomeController {

        public header: string = "Welcome";
        public showNotification = false;
        private notification: INotification;

        static $inject = ['$scope'];

        constructor(private $scope: ng.IScope) {
        }

        // Reads data from current document selection and displays a notification
        public getDataFromSelection(): void {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                (result: Office.AsyncResult): void => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        this.setNotification('The selected text is:', '"' + result.value + '"');
                    } else {
                        this.setNotification('Error:', result.error.message);
                    }
                    // seems we need to explicity apply scope when changing state from Office callback
                    this.$scope.$apply();
                }
            );
        }

        public hideNotification(): void {
            this.showNotification = false;
        }

        // After initialization, expose a common notification function
        private setNotification(header:string, body: string): void {
            this.notification = { header: header, body: body };
            this.showNotification = true;
        };

    }

    angular.module("app")
        .controller("app.HomeController", HomeController);
}