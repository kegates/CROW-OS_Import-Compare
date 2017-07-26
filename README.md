# CROW-OS_Import-Compare
Modifies CROW Transmission Outages Document based on pulled (already inserted into EMS) data from OS_Import Transmission Outages.  This work was done for ISO New England


=============================================================================================
@filename: README.txt
@author: Kevin Gates
@version: 1.0

@date: 07/11/2017 11:28:00

@description: Contains information on the current version of CROW compare.
	      Will be updated as program is updated. (Description, Instructions, and Instillation)
=============================================================================================

DESCRIPTION: This addin is a set of macros meant to compare crow transmission outages
	     to OS_Import data already pulled into EMS.  It can take a csv file (from
	     OS_Import originally) and generate a list of IDs.  It can then use this
	     list of IDs to compare with a CROW transmission outage excel doc.  All outages
	     in OS_Import will be approved in the CROW doc (with a + sign) and all 
	     non-included outages will either be labeled NIC or not at all (meaning the 
	     macro is unsure). The user can then go back using another macro and 
	     change this for any of the outages.  Finally, a faceplate can be made from 
	     the excel doc containing only approved outages.

---------------------------------------------------------------------------------------------

INSTRUCTIONS: 
	      1. If haven't already, set default printer to Generic / Text Only
		(Start Menu -> Type Devices and Printers -> Right Click the Printer
		 and select 'Select as Default Printer')

	      2. Bring in Basecase and go to OS_Import pulling in data for the
		 date under study.  Approve the outages (optional)

	      3. Go to Navigate -> Print -> Portrait (non-color).  You will be
		 asked to specificy a file path.

	      4. Write in the file path and press 'Enter'.  The file path can be 
		 anything you wish so long as it's valid.  Make sure you end the
		 filepath with '<filename>.csv' where <filename> is replaced by
		 whatever name you want.  An example can be seen bellow
		 (ex. C:\Users\exampleUserName\Desktop\exampleFilename.csv)
		 You can now set your default printer back to normal.
	
	      5. Go to CROW -> Reports -> Transmission -> ISO Report, select
		 the day under study for 'Start on date or before' as well as
		 for 'End on date or after'.  Make sure 'Include Outage Periods'
		 is ticked 'Yes'
	
	      6. You may use the CROW add-in the highlight the report or use
		 filters -> (P) Summary Lines.  Other applications of the CROW
		 add-in may work but are not officially supported.

	      7. In the Add-Ins Ribbon, go to the compare dropdown list and click
		 'get OS_Import .csv'.  A file explorer will open up, use it to find
		 and select the csv file you save in step 4.

	      8.  Go to Add-Ins -> Compare -> Compare which will compare the two files
		  and update the excel doc.  You can print it out if you want and complete
		  undecided cases by hand or proceeed to step 9.

	      (Steps 9 and onward are optional)

	      9. You may use any macro in Add-Ins -> Compare -> Step 3: Optional.
		 Their Functionality is listed bellow:

			Decide on Unsures: Will take user through each undecided case
					   and ask if it should be in the case.  If you click
					   'yes' make sure to ADD THAT OUTAGE INTO THE CASE 
					   manually because it is not there already.

			Modify Outage: Lets user change an outage to included, NIC, or Unsure
				       by hand.

			Check Outage: Lets user check if an outage is included, NIC, or Unsure

			Generate Faceplate: Generates a faceplate much like the CROW Add-In except
					    that only outages labelled included are put into the faceplate
					    so that minimal updating is needed.	
    
---------------------------------------------------------------------------------------------

INSTILLATION:
		1. Open the zipped folder and move the folder inside your Microsoft AddIns
		   folder.  The Add-In will not work if you do not move the folder here!

			Example: C:\Users\<username>\AppData\Roaming\Microsoft\AddIns

		2. Open Excel and go to File -> Options -> Add-Ins -> Go... (Next to Manage Dropdown)
		   NOTE: Make sure 'Manage:' is set to 'Excel Add-Ins' before clcking 'Go...'

		3. Click Browse and find the folder you moved in step 1.  Within that folder, select the
		   file ending in '.xlam', it should be called 'CROW comp'.  Make sure 'CROW Comp' Appears
		   checked in the 'Add-Ins available' list.

		4. Open COMPInstallToMenu.xlsm from the folder and click 'Install Compare Menu', it should
		   appear under or above the CROW Add-In in the Add-Ins Tab.  If not, contact Kevin Gates
		   at short term outage (Supervisor: Norman Sproehnle)

		5. To create the generic text printer open the start menu and type/select 'Devices and Printers'
		   Right click a white space on the page that opens and click 'Add a Printer'

		6. In the Menu, click 'Add a local printer' and click 'Next'

		7. In the Next Menu, make sure 'Use an existing port' is ticked.  Next to that, click the dropdown
		   menu and select 'FILE: (Print to File)' and click 'Next'

		8. In the Next Menu, under 'Manufacturer go down to 'Generic' and select it.  Then in the
		   menu to the right labaled 'Printers' select 'Generic / Text Only' and click 'Next'

		9. In the Next Menu, make sure 'Use the driver that is currently installed(recommended)'
		   is ticked and click 'Next'
	
		10. In the Next Menu, name the printer whatever you want ('Generic / Text Only' is 
		    recommended) and click 'Next'

		11. In the Next Menu, make sure 'Do not share this printer' is ticked and click 'Next'
		
		12. Click 'Finish' on the next page

