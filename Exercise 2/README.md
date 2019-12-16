# Exercise 2: Interact with the Excel Workbook using Office.js #
In Exercise 2, you will modify the Office add-in to inject a table of data into the Excel workbook. If you get lost or stuck in the exercise, you can find a completed solution of Exercise 2 at [http://aka.ms/M1s0sh](http://aka.ms/M1s0sh "http://aka.ms/M1s0sh") and a full video walk-though at [https://www.youtube.com/watch?v=1vHNsCDy3wQ](https://www.youtube.com/watch?v=1vHNsCDy3wQ "https://www.youtube.com/watch?v=1vHNsCDy3wQ")

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
2. Next, turn your attention to layout. Open **Home.html** and remove everything inside the _second_ **&lt;div class="padding"&gt;** element and replace it with the following:

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

**[<< Back to home](https://github.com/OfficeDev/Salesforce-Addin-Workshop)**
