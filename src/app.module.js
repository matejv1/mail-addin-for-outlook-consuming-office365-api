(function () {
  'use strict';

  // create the angular app
  var outlookApp = angular.module('appowa', [
    'ngRoute',
    'ngAnimate',
    'ngSanitize',
    'ngMaterial',
    'AdalAngular'
  ]);

  // configure the app
  outlookApp.config(['$logProvider', '$mdThemingProvider', function ($logProvider, $mdThemingProvider) {
    // set debug logging to on
    if ($logProvider.debugEnabled) {
      $logProvider.debugEnabled(true);
    }

    // configure theme color
    $mdThemingProvider.theme('default')
                      .primaryPalette('blue');
  }]);

  // when office has initalized, manually bootstrap the app
  

})();