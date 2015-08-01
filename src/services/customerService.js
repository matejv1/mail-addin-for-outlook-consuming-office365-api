(function () {
  'use strict';

  angular.module('appowa')
      .service('customerService', ['$q', '$http', customerService]);

  /**
   * Custom Angular service that talks to a static JSON file simulating a REST API.
   */
  function customerService($q, $http) {
    // public signature of the service
    return {
      lookupCustomerPartials: lookupCustomerPartials,
      lookupCustomer: lookupCustomer,
      getRelatedDocs: getRelatedDocs,
      getFiles: getFiles

    };

    /** *********************************************************** */

    /**
     * Queries the remote service for possible customer matches.
     *
     * @param possibleCustomers {Array<string>}   Collection of customer last names to lookup.
     */
    function lookupCustomerPartials(possibleCustomers) {
      var deferred = $q.defer();

      // if nothing submitted return empty collection
      if (!possibleCustomers || possibleCustomers.length == 0) {
        deferred.resolve([]);
      }

      // fetch data
      var endpoint = '/content/customers.json';

      // execute query
      $http({
        method: 'GET',
        url: endpoint
      }).success(function (response) {
        var customers = [];

        // look at each customer to find a match
        response.d.results.forEach(function (customer) {
          if (possibleCustomers.indexOf(customer.LastName) != -1) {
            customers.push(customer);
          }
        });

        deferred.resolve(customers);
      }).error(function (error) {
        deferred.reject(error);
      });

      return deferred.promise;
    }

    /**
     * Finds a specific customer form the datasource.
     *
     * @param customerID  {number}    Unique ID of the customer.
     */
    function lookupCustomer(customerID) {
      var deferred = $q.defer();

      // fetch data
      var endpoint = '/content/customers.json';

      $http({
        method: 'GET',
        url: endpoint
      }).success(function (response) {
        var result = {};

        // find the matching customer
        response.d.results.forEach(function (customer) {
          if (customerID == customer.CustomerID) {
            result = customer;
          }
        });

        deferred.resolve(result);
      }).error(function (error) {
        deferred.reject(error);
      });

      return deferred.promise;
    }

    function getRelatedDocs(mail) {
      var deferred = $q.defer();
      deferred.resolve(mail);
      // fetch data
      console.log(mail.from.emailAddress);

      return deferred.promise;
    }

    function getFiles(mailbox) {
      var deferred = $q.defer();
      console.log(mailbox.from.emailAddress);
      // fetch data
      var restQueryUrl = "https://agile9.sharepoint.com/_api/search/query?querytext='" + mailbox.from.emailAddress + "'";

      $http({
        method: 'GET',
        url: restQueryUrl,
        headers: {
            "accept": "application/json; odata=verbose",
        }
      }).success(function (data) {
        var result = {};

        // find the matching customer
        result = $.map(data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results, function (item) {
                return getFields(item.Cells.results);
            });

        deferred.resolve(result);
      }).error(function (error) {
        deferred.reject(error);
      });

      return deferred.promise;
    }

  }
})();

//helper function for rest-search result formating
function getFields(results) {
    r = {};
    for (var i = 0; i < results.length; i++) {
        if (results[i] != undefined && results[i].Key != undefined) {
            r[results[i].Key] = results[i].Value;
        }
    }
    return r;
}
