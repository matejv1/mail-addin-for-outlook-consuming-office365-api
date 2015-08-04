(function () {
  'use strict';

  var outlookApp = angular.module('appowa');

  // load routes
  outlookApp.config(['$routeProvider', '$httpProvider', 'adalAuthenticationServiceProvider', routeConfigurator]);

  function routeConfigurator($routeProvider, $httpProvider, adalProvider) {

    //Initialize ADAL
    adalProvider.init({
        tenant: "agile9.onmicrosoft.com",
        clientId: "9e03550c-1678-4093-9b12-05946b4df46b",
        cacheLocation: "localStorage",
        endpoints: {
            'https://agile9.sharepoint.com/_api/': 'https://agile9.sharepoint.com',
            'https://agile9-my.sharepoint.com/_api/v1.0/me': 'https://agile9-my.sharepoint.com',
            'https://outlook.office365.com/api/v1.0/me': 'https://outlook.office365.com'
        }
    }, $httpProvider);
    
    $routeProvider
        .when('/', {
          templateUrl: '/views/home-view.html',
          controller: 'homeController',
          requireADLogin: true
        })
        .when('/files', {
          templateUrl: '/views/files-view.html',
          controller: 'homeController',
          controllerAs: 'vm',
          requireADLogin: true
        })
        .when('/mails', {
          templateUrl: '/views/mails-view.html',
          requireADLogin: true
        })
        .when('/employees', {
          templateUrl: '/views/employees-view.html',
          requireADLogin: true
        })
        .when('/reports', {
          templateUrl: '/views/reports-view.html',
          requireADLogin: true
        });
    $routeProvider.otherwise({redirectTo: '/'});
  }
})();