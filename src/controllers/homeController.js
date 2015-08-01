(function () {
  'use strict';

  angular.module('appowa')
      .controller('homeController',
      ['$q', '$location', 'officeService', 'customerService',
        homeController]);

  /**
   * Controller constructor
   * @param $q                Angular's $q promise service.
   * @param $location         Angular's $location service.
   * @param officeService     Custom Angular service for talking to the Office client.
   * @param customerService   Custom Angular service for customer data.
   */
  function homeController($q, $location, officeService, customerService) {
    var vm = this;


    /** *********************************************************** */

    Office.initialize = function () {
      console.log(">>> Office.initialize()");
      
      init();
    };

    /**
     * Initialize the controller
     */
    function init() {
      getCurrentMailboxItem()
          .then(function(){
            return getFiles();
          });
    }
    
    function getCurrentMailboxItem(){
      var deferred = $q.defer();

      officeService.getCurrentMailboxItem()
          .then(function(mailbox){

            vm.currentMailboxItem = mailbox;
            deferred.resolve();
          })
          .catch(function (error) {
              deferred.reject(error);
          });

      return deferred.promise;
    }

    function getFiles(){
      var deferred = $q.defer();

      customerService.getFiles(vm.currentMailboxItem)
          .then(function(files){

            vm.files = files;
            deferred.resolve();
          })
          .catch(function (error) {
              deferred.reject(error);
          });

      return deferred.promise;
    }

  }

})();