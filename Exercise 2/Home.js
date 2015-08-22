/// <reference path="../App.js" />
// global app

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
			$('h1').html('Export Options');
			
			//toggle filter options
			$('.option').click(function() {
				if ($(this).hasClass('active'))
					$(this).removeClass('active');
				else
					$(this).addClass('active');
			});
			
			//hard-code the table headers for opportunities
			var fields = ['Id', 'Account.Name', 'Name', 'Description', 'StageName', 'Amount', 'Probability', 'ExpectedRevenue', 'TotalOpportunityQuantity', 'CloseDate', 'Type'];
			
			//sign-in and get data from Salesforce
            $('#btnExport').click(function() {
				//initialize headers
				var headers = new Array();
				$(fields).each(function(field_index, field) {
					headers.push(field);
				});
				
				//initialize the Office Table
				var officeTable = new Office.TableData();
				officeTable.headers = headers;
				
				//use hard-coded random data...exercise 3 will convert this to live data
				officeTable.rows = [
					[ '001', 'Account 1', 'Some Opportunity 1', 'Some Description 1', 'Qualification', 123, 50, 123, 1, '5/5/2016', 'New Customer' ],
					[ '002', 'Account 2', 'Some Opportunity 2', 'Some Description 2', 'Qualification', 123, 50, 123, 1, '5/5/2016', 'New Customer' ]
				];
				
				//insert the data into Excel workbook by calling setSelectedDataAsync
				Office.context.document.setSelectedDataAsync(officeTable, 
					{
						coercionType: Office.CoercionType.Table, 
						cellFormat: [ { cells: Office.Table.All, format: { width: "auto fit" } } ] 
					}, function (result) {
					if (result.status == Office.AsyncResultStatus.Succeeded)
						app.showNotification("SUCCESS", "Opportunities loaded: " + officeTable.rows.length);
					else
						app.showNotification("ERROR", "Writing opportunities to Excel failed");
				});
			});
        });
    };
})();