# Building Office add-ins with HTML5, JavaScript, and Salesforce.com #

In this workshop, you will build a basic Office add-in, modify it to interact with an Excel workbook, and connect it to the Salesforce.com APIs. Modern Office add-ins are built using web technologies such as HTML5, JavaScript, and CSS. The exercises in this workshop will use an in-browser web editor called Napa. All you will need to complete the workshop is a modern browser and a few accounts (see ***Prerequisites*** section).

## Prerequisites ##
1. You must have a Microsoft account (an email address and password that you use to sign in to all Microsoft sites and services, including Outlook.com, Xbox Live, OneDrive, and Office 365). If you don't have one, you can quickly create one at [https://signup.live.com/signup](https://signup.live.com/signup "https://signup.live.com/signup")
2. You must have a Salesforce.com developer account to complete Exercise 3. If you don't have one, you can quickly create one at [https://developer.salesforce.com/signup](https://developer.salesforce.com/signup "https://developer.salesforce.com/signup")

## Exercise 1: Create an Office Add-in using Napa ##
In Exercise 1 you will create your first Office Add-in using Napa. More specifically, you will create a task pane add-in for Microsoft Excel. In subsequent exercises you will add additional functionality to the add-in, including integration with Salesforce APIs. If you get lost or stuck in the exercise, you can find a completed solution of Exercise 1 at [http://aka.ms/Qeqq0j](http://aka.ms/Qeqq0j "http://aka.ms/Qeqq0j")

1. Open a browser and navigate to Napa ([https://www.napacloudapp.com/](https://www.napacloudapp.com/ "https://www.napacloudapp.com/")).
2. Click the **Sign in** link in the upper right and sign-in with a Microsoft account.
![Sign-in](http://i.imgur.com/IpnTOaX.png)
3. After completing sign-in, Napa will display a screen asking **What type of Add-in do you want to build?** with options to build **Content**, **Task Pane**, and **Mail** Add-ins for Office.
4. Select **Task pane Add-in for Office**, give the project a name (ex: Opportunity Quick Editor), and click Create.
![Select add-in type](http://i.imgur.com/tsWjgW7.png)
5. When Napa finishes creating the project, it will bring you into the editor window. Take some time to get acquainted with the Napa editor and the starter project.
6. The toolbar on the far left contains actions and commands for the working in Napa, including commands for file explorer, properties, run project, retract add-in, share project, publish, and more.
7. The starter project already contains a number of HTML, JavaScript, and CSS files, including **Home.html** which should already be loaded in the editor pane and is the primary page for the add-in user interface.
8. Notice that **Home.html** already references **JQuery** and **Office.js**, which provides the integration with Office documents.
![Home.html](http://i.imgur.com/5ckxf8B.png)
9. Open **Home.js** and add **$('h1').html('Export Options');** to **line 11**. Notice how the Napa editor provides script help/hints as you type (often called "**IntelliSense**").
![IntelliSense](http://i.imgur.com/7nIh0Jr.png)
10. Click the **Run Project** button (the play button) in the left toolbar to run the add-in in Excel Online. If the browser prevents the Excel Online window from opening, you can manually launch it by clicking the **Click here to launch your app in a new window** link.
11. Since this is the first time using the add-in, Excel Online may prompt you to start the add-in manually. Click the **START** button in the task pane to allow the add-in to load.
![Prompt to Start](http://i.imgur.com/9tvINlz.png)
12. After the task pane add-in loads, you need to toggle the workbook into edit mode by select **Edit Workbook** > **Edit in Excel Online** from the menu in the header.
![Edit mode](http://i.imgur.com/0FWnH1B.png)
13. Type some text into any of the workbook cells (ex: "Hello Dreamforce") and then click the **Get data from selection** button in the task pane add-in. Notice that the add-in can read data from the document. Later, you will write data to the document from the add-in.
![getSelectedDataAsync](http://i.imgur.com/zqMP6zp.png)
14. Also notice how the add-in header (blue section) says "***Export Options***", which you set via script in step #9.
15. Congratulations, you have written your first Office add-in!!!
## Exercise 2: Interact with the document using Office.js ##
In Exercise 2, you will modify the Office add-in to inject a table of data into the Excel workbook. If you get lost or stuck in the exercise, you can find a completed solution of Exercise 2 at [http://aka.ms/M1s0sh](http://aka.ms/M1s0sh "http://aka.ms/M1s0sh")

1. First, you will define some style classes for the add-in. Open **Home.css** and add the following classes:

	    /* Page-specific styling */
		.option {
			padding: 6px;
			border-top: 1px dotted #CCC;
			cursor: pointer;
		}
		
		.option:last-child {
			border-bottom: 1px dotted #CCC;
		}
		
		.option:hover {
			background-color: #DFEEF9;
		}
		
		.option.active {
			background-color: #BFDDF2
		}
2. Next, turn your attention to layout. Open **Home.html** and remove everything inside the **&lt;div class="padding"&gt;** element and replace it with the following:

		<p>Please select the opportunity stage(s) you want to export:</p>
		<div>
            <div class="option">Prospecting</div>
			<div class="option">Qualification</div>
			<div class="option">Needs Analysis</div>
			<div class="option">Value Proposition</div>
			<div class="option">Id. Decision Makers</div>
			<div class="option">Perception Analysis</div>
			<div class="option">Proposal/Price Quote</div>
			<div class="option">Negociation/Review</div>
			<div class="option">Closed Won</div>
			<div class="option">Closed Lost</div>
		</div>
		<br/>
		<button id="btnExport">Export to Excel</button>
3. Finally, modify the functionality of the add-in by updating **Home.js**. Start by removing the **getDataFromSelection** function in **lines 16-27** that will no longer be used.
4. Next, replace the button click script in **line 12** with script to toggle the stage type selections:

		//toggle filter options
		$('.option').click(function() {
			if ($(this).hasClass('active'))
				$(this).removeClass('active');
			else
				$(this).addClass('active');
		});
	
5. Below the stage type toggle script, define a string array that represents the headers of the opportunities table:

		//hard-code the table headers for opportunities
		var fields = ['Id', 'Account.Name', 'Name', 'Description', 'StageName', 'Amount', 'Probability', 
			'ExpectedRevenue', 'TotalOpportunityQuantity', 'CloseDate', 'Type'];
6. Finally, create a new click event on the **btnExport** button that initializes a **Office.TableData** object, loads it with headers and some hard-coded rows, and writes it to the Excel workbook using the **Office.context.document.setSelectedDataAsync** function:

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
7. The completed **Home.js** file should look like the following:

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
8. Click the **Run Project** button (the play button) in the left toolbar to run the add-in in Excel Online. If the browser prevents the Excel Online window from opening, you can manually launch it by clicking the **Click here to launch your app in a new window** link.
9. Click on some of the stage type options and notice how the row highlighting toggles. This exercise set up the toggle events, but you won't actually use the selections until Exercise 3.
10. Click on the **Export to Excel** button to write the hard-coded data into the Excel workbook as a new table.
![setSelectedDataAsync](http://i.imgur.com/6OENiGu.png)
11. Congratulations, you have used your Office add-in to write complex data into an Excel Workbook!!!
## Exercise 3: Connect to Salesforce.com APIs ##
In Exercise 3, you will modify the Office add-in to sign-in and query data from Salesforce.com. If you get lost or stuck in the exercise, you can find a completed solution of Exercise 3 at [http://aka.ms/X80ts0](http://aka.ms/X80ts0 "http://aka.ms/X80ts0")

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

## Going Further (optional) ##
Want to take this further? Try modifying the add-in to highlight cells as they are updated. You can do this by adding table binding and then using an event handler for **BindingDataChanged**. The [Excel API Tutorial](https://store.office.com/api-tutorial-content-WA104077907.aspx "Excel API Tutorial") shows how to accomplish both of these tasks. You can also reference the completed solution at [http://aka.ms/Gs5kib](http://aka.ms/Gs5kib "http://aka.ms/Gs5kib")
![Change formatting](http://i.imgur.com/kY4Pckg.png)