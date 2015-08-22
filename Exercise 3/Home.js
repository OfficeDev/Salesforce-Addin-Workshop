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
            $('#btnExport').sfLoginButton(function(token) {
				//check that the token isn't null
				if (token !== null) {
					//build select fields
					var select_fields = '';
					var headers = new Array();
					$(fields).each(function(field_index, field) {
						if (field_index > 0)
							select_fields += ', ';
						select_fields += field;
						headers.push(field);
					});
				
					//build up a where filter based on stage selections
					var where = '';
					$('.active').each(function(filter_index, filter) {
						if (filter_index > 0)
							where += ' or ';
						where += 'StageName = \'' + $(filter).text() + '\'';
					});
					
					//build the query
					var query = encodeURIComponent('select {0} from Opportunity where {1}'.replace('{0}', select_fields).replace('{1}', where));
				
					//initialize the Office Table
					var officeTable = new Office.TableData();
					officeTable.headers = headers;
				
					//execute the REST query against Salesforce
					$.ajax({
						url: token.instance_url + "/services/data/v20.0/query/?q=" + query,
						headers: {
							"Authorization": "Bearer " + token.access_token,
							"accept": "application/json;odata=verbose"	
						},
						success: function (data) {
							//loop through the returned records and append to the officeTable
							$(data.records).each(function(row_index, row) {
								var data = new Array();
								$(fields).each(function(field_index, field) {
									if (field === 'Account.Name')
										data.push(row.Account.Name);
									else
										data.push(row[field]);
								});
								officeTable.rows.push(data);
							});
							
							//insert the data into Excel by calling setSelectedDataAsync
							Office.context.document.setSelectedDataAsync(officeTable, 
								{
									coercionType: Office.CoercionType.Table, 
									cellFormat: [ { cells: Office.Table.All, format: { width: "auto fit" } } ] 
								}, function (result) {
								if (result.status == Office.AsyncResultStatus.Succeeded)
									app.showNotification("SUCCESS", "Opportunities loaded: " + data.records.length);
								else
									app.showNotification("ERROR", "Writing opportunities to Excel failed");
							});
						},
						error: function (err) {
							app.showNotification('ERROR:', 'Opportunities failed to load');
						}
					});
				}
				else
					app.showNotification('Error', 'Error establishing connection to Salesforce');
			});
        });
    };
})();