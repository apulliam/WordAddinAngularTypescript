/// <reference path="../../typings/tsd.d.ts" />
var WordAddIn;
(function (WordAddIn) {
    var HomeController = (function () {
        function HomeController($scope) {
            this.$scope = $scope;
            this.header = "Welcome";
            this.showNotification = false;
        }
        // Reads data from current document selection and displays a notification
        HomeController.prototype.getDataFromSelection = function () {
            var _this = this;
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    _this.setNotification('The selected text is:', '"' + result.value + '"');
                }
                else {
                    _this.setNotification('Error:', result.error.message);
                }
                // seems we need to explicity apply scope when changing state from Office callback
                _this.$scope.$apply();
            });
        };
        HomeController.prototype.hideNotification = function () {
            this.showNotification = false;
        };
        // After initialization, expose a common notification function
        HomeController.prototype.setNotification = function (header, body) {
            this.notification = { header: header, body: body };
            this.showNotification = true;
        };
        ;
        HomeController.$inject = ['$scope'];
        return HomeController;
    })();
    angular.module("app")
        .controller("app.HomeController", HomeController);
})(WordAddIn || (WordAddIn = {}));
//# sourceMappingURL=home.controller.js.map