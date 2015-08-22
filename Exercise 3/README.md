# Exercise 3: Connect to Salesforce.com APIs #
In Exercise 3, you will modify the Office add-in to sign-in and query data from Salesforce.com. If you get lost or stuck in the exercise, you can find a completed solution of Exercise 3 at [http://aka.ms/X80ts0](http://aka.ms/X80ts0 "http://aka.ms/X80ts0") and a full video walk-though at [https://www.youtube.com/watch?v=_lXp3ML0W3k](https://www.youtube.com/watch?v=_lXp3ML0W3k "https://www.youtube.com/watch?v=_lXp3ML0W3k")

> **NOTE**: Napa does not currently expose reply URLs, which are required to complete a web server OAuth flow with Salesforce. To work around this limitation, the exercise will leverage a hosted proxy to help complete OAuth flows with Salesforce. For more information on a typical OAuth flow, see Salesforce documentation on [Understanding the Web Server OAuth Authentication Flow](https://developer.salesforce.com/docs/atlas.en-us.api_rest.meta/api_rest/intro_understanding_web_server_oauth_flow.htm "Understanding the Web Server OAuth Authentication Flow"). For more information on the hosted proxy, you can view the full project on GitHub ([https://github.com/richdizz/Salesforce-Proxy](https://github.com/richdizz/Salesforce-Proxy "https://github.com/richdizz/Salesforce-Proxy")).

1. First, add a script reference to **https://o365workshop.azurewebsites.net/scripts/sfproxy.js** in the head of **Home.html** (add it directly after the office.js reference):

		<script src="https://o365workshop.azurewebsites.net/scripts/sfproxy.js" type="text/javascript"></script>
2. Next, you need to convert the **btnExport** button to authenticate against Salesforce before performing additional actions. To do this, change the btnExport click event from this:

		$('#btnExport').click(function() {
			...omitted for brevity
To the **sfLoginButton** extension seen below. Also note that the **sfLoginButton** callback provides a token you will include on API calls to Salesforce:

		$('#btnExport').sfLoginButton(function(token) {
			...omitted for brevity
3. You should validate the token returned isn't null and move any additional logic inside this check.

		//check that the token isn't null
		if (token !== null) {
			...additional logic omitted for brevity
		}
		else
			app.showNotification('Error', 'Error establishing connection to Salesforce');
4. Next, you need to build up a query using Salesforce Object Query Language (SOQL). Specifically, you need to build a string of select fields and where clauses. To start, modify the initialize headers code as follows:

		//build select fields
		var select_fields = '';
		var headers = new Array();
		$(fields).each(function(field_index, field) {
			if (field_index > 0)
				select_fields += ', ';
			select_fields += field;
			headers.push(field);
		});

5. Next, you need to build up a where clause based on the stage type selections. The stage type toggle logic was implemented in **Step 4** of **Exercise 2**, but here is where stage types come to life as filters in your add-in.
				
		//build up a where filter based on stage selections
		var where = '';
		$('.active').each(function(filter_index, filter) {
			if (filter_index > 0)
				where += ' or ';
			where += 'StageName = \'' + $(filter).text() + '\'';
		});
6. Combine the **select_fields** and **where** clause into a single SOQL statement.

		//build the query
		var query = encodeURIComponent('select {0} from Opportunity where {1}'.replace('{0}', select_fields).replace('{1}', where));
8. Next, find and replace the script that hard-codes **officeTable.rows** with a JQuery call to the Salesforce APIs. Notice **token.access_token** is passed in the **Authorization** header of the REST call. Without this, Salesforce would deny the add-in access to data.

		//execute the REST query against Salesforce
		$.ajax({
			url: token.instance_url + "/services/data/v20.0/query/?q=" + query,
			headers: {
				"Authorization": "Bearer " + token.access_token,
				"accept": "application/json;odata=verbose"	
			},
			success: function (data) {
				//TODO: loop through data and write to Excel
			},
			error: function (err) {
				app.showNotification('ERROR:', 'Opportunities failed to load');
			}
		});
9. Finally, you can add script to loop through the returned JSON, add the data to the **officeTable.rows**, and then call **setSelectedDataAsync** code similar to Exercise 1. The completed **Home.js** should look like this:

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
10. Click the **Run Project** button (the play button) in the left toolbar to run the add-in in Excel Online. If the browser prevents the Excel Online window from opening, you can manually launch it by clicking the **Click here to launch your app in a new window** link.
10. When the add-in loads, the **Export to Excel** button might initially be disabled while scripts finish loading. When it finally enables, select a few stage types and click the export button.
11. Clicking the export button will launch an authentication dialog. Complete a sign-in to Salesforce and allow the app access to Salesforce if prompted:
![sign-in](http://i.imgur.com/K7uO0mH.png)
12. After authenticating and granting the app access to Salesforce, the dialog will close and the add-in will write opportunity records from Salesforce into the Excel workbook:
![Opportunity Export](http://i.imgur.com/axyfSRl.png)
13. Congratulations, you have built a fully functioning Office add-in for Excel that connects to Salesforce!!!

**[<< Back to home](https://github.com/OfficeDev/Salesforce-Addin-Workshop)**