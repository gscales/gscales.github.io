﻿var o365CorsApp = angular.module("o365CorsApp", ['ngRoute', 'AdalAngular'])
o365CorsApp.factory("ShareData", function () {
    return { value: 0 }
});
o365CorsApp.config(['$routeProvider', '$httpProvider', 'adalAuthenticationServiceProvider', function ($routeProvider, $httpProvider, adalProvider) {
    $routeProvider
           .when('/',
           {
               controller: 'ContactsController',
               templateUrl: 'Contacts.html',
               requireADLogin: true
           })
           .otherwise({ redirectTo: '/' });

    var adalConfig = {
        tenant: '1c3a18bf-da31-4f6c-a404-2c06c9cf5ae4',
        clientId: '11eb2c22-13fe-44e0-89c0-e873d589e2d0',
        extraQueryParameter: 'nux=1',
        endpoints: {
           "https://outlook.office365.com/api/v1.0": "https://outlook.office365.com/"
        }
    };
    adalProvider.init(adalConfig, $httpProvider);
}]);
o365CorsApp.controller("ContactsController", function ($scope, $q, $location, $http, ShareData, o365CorsFactory) {
    o365CorsFactory.getContacts().then(function (response) {
        $scope.contacts = response.data.value;
    });

});

o365CorsApp.factory('o365CorsFactory', ['$http' ,function ($http) {
    var factory = {};
   
    factory.getContacts = function () {
        return $http.get('https://outlook.office365.com/api/v1.0/me/contacts')
    }

    factory.getContact = function (id) {
        return $http.get('https://outlook.office365.com/api/v1.0/me/contacts/'+id)
    }

    return factory;
}]);

















