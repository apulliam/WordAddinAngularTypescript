/// <reference path="../typings/tsd.d.ts" />
((): void => {


    config.$inject = ["$routeProvider"];

function config($routeProvider: ng.route.IRouteProvider) {
    // Configure the routes.
    $routeProvider
        .when("/home", {
            templateUrl: "/app/home/home.html",
            controller: "app.HomeController",
            controllerAs: "vm"
        })
        .when("/demo", {
            templateUrl: "/app/demo/demo.html",
            controller: "app.DemoController",
            controllerAs: "vm"
        })
        .otherwise({
            redirectTo: "/home"
        });
}

    angular
        .module("app").
        config(config);
})();