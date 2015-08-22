# Exercise 1: Create an Office Add-in using Napa #
In Exercise 1 you will create your first Office Add-in using Napa. More specifically, you will create a task pane add-in for Microsoft Excel. In subsequent exercises you will add additional functionality to the add-in, including integration with Salesforce APIs. If you get lost or stuck in the exercise, you can find a completed solution of Exercise 1 at [http://aka.ms/Qeqq0j](http://aka.ms/Qeqq0j "http://aka.ms/Qeqq0j") and a full video walk-though at [https://www.youtube.com/watch?v=skvou346HOo](https://www.youtube.com/watch?v=skvou346HOo "https://www.youtube.com/watch?v=skvou346HOo")

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

**[<< Back to home](https://github.com/OfficeDev/Salesforce-Addin-Workshop)**