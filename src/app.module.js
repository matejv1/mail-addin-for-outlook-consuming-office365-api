(function () {
  'use strict';

  // create the angular app
  var outlookApp = angular.module('appowa', [
    'ngRoute',
    'ui.bootstrap',
    'AdalAngular'
  ]);

  // configure the app
  outlookApp.config(['$logProvider', function ($logProvider) {
    // set debug logging to on
    if ($logProvider.debugEnabled) {
      $logProvider.debugEnabled(true);
    }

  }]);

  // when office has initalized, manually bootstrap the app
  

})();