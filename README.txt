MobileView Assistant
Developed by Ryan Mooney, BMET II, at Mercy Medical Center - Cedar Rapids, IA

This application is meant to work with MobileView software in order to streamline and organize
asset data into an easy to read spreadsheet that can be used to keep track of location data for
your institutions assets.

Prerequisites:
-Access to MobileView software through Internet Explorer
-A user (or admin) account for MobileView
-Microsoft Excel (to view results)

Prerequisites if running the Python, non-compiled version:
-Python libraries: tkinter, selenium, IEDriverSoftware, xlwt, xlrd, smtplib

Installation:
1. Log in to your admin account (if necessary).
2. Copy and paste the whole MobileViewAssistant folder into your C: drive
3. Open up the MobileViewAssistant folder and find the directory path for IEDriverSoftware.exe
	-This is found by right clicking IEDriverSoftware.exe, click properties, and going to location:
	-Mine was C:\MobileViewAssistant
4. Add this path plus the name ("IEDriverSoftware.exe") to your computers PATH (do this as admin and regular user)
	-Search Environment Variables in the Search Bar
	-Click "Edit the system Environment Variables"
	-Click "Environment Variables"
	-Under User Variables, select Path, then click Edit
	-Click New on the right
	-Add the drivers path plus its name(ie, "C:\MobileViewAssistant\IEDriverSoftware.exe")
	-Click Ok until you are out.
5. Now, we need to make sure Internet Explorers settings are correct. As an admin, open up Internet Explorer
	and click the settings gear in the upper right.
6. Click Security.
7. Make sure that for each type of site, Protected Mode is either enabled or disabled for each (ie, all are enabled,
	or all are disabled).
8.Done! Exit out of admin and log back in as a normal user.

Setting up Mailing List:
1. To set up the mailing list, open the ResultsMailingList document in the main folder.
2. Add the emails you'd like to send the results to (if desired), seperated by semicolons (";").

Opening the Software:
1. Go to the MobileViewAssistant main folder.
2. Hold shift and right click on the MVAssistant Application icon in the main folder. Select "Run as other user" and
	sign in with your admin account.
3. The application should now open normally.

Testing the software:
1. Once open, the default settings will be selected.
2. Select "Run a test trial only."
3. If the IEDriverSoftware path is set up properly, a test case should run, save, and then a spreadsheet should open.
4. If this does not happen, check that the driver path is in your systems PATH, and check that Microsoft Excel is installed.

Running the software:
1. To run a normal trial, open the software using Shift+Right click, selecting "Run as other user" and log in with your admin
	credentials.
2. Once open, enter your NORMAL username and password that you use to log in to Mobile View. 
	-Unless you have a special asset list you want to run, leave the asset file as "DefaultAssetList". 
	-Select if you want to run the trial on "All Assets" or	just for a certain PM Month. 
	-If you have access to Mobile View's admin pages, leave "I have access to Mobile View Admin" selected.
		This is the quickest way to find the asset locations. If you don't have access, uncheck this box.
		The application will then instead use the regular location finder in Mobile View.
	-If you would like to email the results once done, check the corresponding box.
3. Then click "Run". The application will open up Internet Explorer, navigate to Mobile View and use your credentials
	to find the location data for the assets requested. 
4. This will take some time, but the application can do this in	the background. Feel free to continue working while
	it searches!

Creating new asset lists:
1. You can create and use your own asset lists as long as you follow a basic outline.
2. Your list must be created in Excel.
3. In the first row of your Excel spreadsheet, have "Asset".
4. Below "Asset", list the assets you'd like to locate, going down the first column of the spreadsheet