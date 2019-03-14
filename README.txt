MobileView Assistant
Developed by Ryan Mooney, BMET II, at Mercy Medical Center - Cedar Rapids, IA

This application is meant to work with MobileView software in order to streamline and organize asset data into an easy to read spreadsheet that can be used to keep track of location data for your institutions assets.

Prerequisites:
-Access to MobileView software through Internet Explorer
-A user (or admin) account for MobileView
-Admin user account (For Mercy Employees)
-Microsoft Excel (to view results)

Prerequisites if running the Python, non-compiled version:
-Python libraries: tkinter, selenium, IEDriverSoftware, xlwt, xlrd, smtplib

Installation:
1. Log in to your admin account (if necessary).
2. Copy and paste the whole MobileViewAssistant folder into your C: drive (this is so it is available in both Admin and Regular User accounts).
3. Open up the MobileViewAssistant folder and find the directory path for IEDriverSoftware.exe
	a.This is found by right clicking IEDriverSoftware.exe, click properties, and going to "Location:"
	b.Mine was C:\MobileViewAssistant
4. Add these paths to the name ("C:\MobileViewAssistant\IEDriverSoftware.exe", "C:\MobileViewAssistant\chromedriver.exe") and add it to your computers PATH (do this as admin and regular user)
	a.Search Environment Variables in the Search Bar
	b.Click "Edit the system Environment Variables"
	c.Click "Environment Variables"
	d.Under User Variables, select Path, then click Edit
	e.Click New on the right
	f.Add the drivers path plus its name(ie, "C:\MobileViewAssistant\IEDriverSoftware.exe", "C:\MobileViewAssistant\chromedriver.exe")
	g.Click Ok until you are out.
5. Now, we need to make sure Internet Explorers settings are correct. As an admin, open up Internet Explorer and click the settings gear in the upper right.
6. Click Security.
7. Make sure that for each type of site, Protected Mode is either enabled or disabled for each (ie, all are enabled, or all are disabled). This is a settings that the Driver needs to run properly.
8.Done! Exit out of admin and log back in as a normal user.

Opening the Software:
1. Go to the C:\MobileViewAssistant main folder (as a normal user).
2. Hold shift and right click on the MVAssistant Application icon in the main folder. Select "Run as other user" and sign in with your admin account.
3. The application should now open normally.

Testing the software:
1. Once open, the default settings will be selected.
2. Select "Run a test trial only."
3. If the paths for the IE and Chrome drivers are set up properly, a test case should run, save, and then a spreadsheet should open.
4. If this does not happen, check that the driver path is in your systems PATH, and check that Microsoft Excel is installed.
NOTE: Sometimes it takes a couple minutes for your computer to recognize the new PATH variables, so wait a few minutes and try again.

Running the software:
1. To run a normal trial, open the software using Shift+Right click, selecting "Run as other user" and log in with your admin credentials.
2. Once open, enter your REGULAR username and password that you use to log in to Mobile View and/or RSQ (if you plan on cross-checking). 
3. Unless you have a special asset list you want to run, leave the asset file as "DefaultAssetList". 
4. Select if you want to run the trial on "All Assets" or just for a certain "PM Month". 
5. If you have access to Mobile View's admin pages, leave "I have access to Mobile View Admin" selected. (This is the quickest way to find the asset locations. If you don't have access, uncheck this box. The application will then instead use the regular location finder in Mobile View.)
	a. Also, if you would like to only print out the assets that currently have PMs, check the "Cross-check" box to use RSQ to cross-check the list.
6. If you would like to email the results once done, check the corresponding box.
7. Then click "Run". The application will open up Internet Explorer, navigate to Mobile View and use your credentials to find the location data for the assets requested. Once done, it will load up the results sheet in Excel.
8. This will take some time, but the application can do this in	the background. Feel free to continue working while it searches!

Setting up Mailing List:
1. To set up the mailing list, open the ResultsMailingList document in the main folder.
2. Add the emails you'd like to send the results to (if desired), seperated by semicolons (";").

Creating new asset lists:
1. You can create and use your own asset lists as long as you follow the template of the default asset list spreadsheet.
2. Your list must be created in Excel.
3. In the first row of your Excel spreadsheet, have "Asset".
4. Below "Asset", list the assets you'd like to locate, going down the first column of the spreadsheet

Compiling the program using pyinstaller:
1. Install pyinstaller from the command line.
2. Go to the folder with the files in it after installing pyinstaller
3. Run: pyinstaller --paths C:\Windows\WinSxS\x86_microsoft-windows-m..namespace-downlevel_31bf3856ad364e35_10.0.17134.1_none_50c6cb8431e7428f --hidden-import tkinter MVAssistantPopupDialog.py