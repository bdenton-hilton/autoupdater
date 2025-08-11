Install:

This will not work and error out on you if you don't have a Hilton device.
Avoid placing the .exe on your desktop, it will make a number of files that will cluter wherever it is.
Unzip / save the EXE anywhere other than One Drive folders. There is no reason this needs to be backed up.
I reccomend saving this somewhere in program data or in C:\Users\USERNAME\AppData\Roaming

To save in the user's appdata, press the windows key + r key at the same time. In the run box type %appdata%, and hit "ok".
make a folder in this directory, name it whatever you want.

Create a shortcut to it/pin it to your start menu if you'd like it to hide off in it's own little folder.

Setup:

If you are extracting this program from a zip file you should have everything you need to make the auto-OnQ reports work.
If you are getting just the EXE however, you will want need to specifically add the import excel module to the PEP Migration folder.
To do so, go into the "helper" folder in the same directory as the EXE. This folder will be made when the program is first launched.
Run the Export ImportExcel.ps1 script to save out the import excel module. Copy the ImportExcel folder.
Go to the PEP Migration folder. Make a folder called "Modules", it is case sensitive. Paste the ImportExcel folder into Modules. You're good to go!

Use:
Enter your ADM credentials. Please note for the scheduled task to run, your password CANNOT be expired
on the day you intend the task to run. If you want to catch a Tuesday migration your password must
be expiring on Wednesday. Don't cut it close, it just won't work.

Select the folder you want to place on the OnQ Server. Most likely it will be the bundled PEP Migration Folder.
You can place any folder with any contents you'd like however. 

Select where you want that folder to be copied to. For the PEP Migration folder, put in "D:\" for the target location.

Enter in any number of inncodes. If for some reason a hotel's OnQ server cannot be reached by inncodeserver.na.hhacpr.hilton.com, you can specify the address yourself. Please do not use an IP address.

If you want to set the task that looks for a day-ran Night Audit to run the OnQ Reports and turn off interfaces, leave the checkbox ticked.

You can export a .csv of the status of the connections at any point.

Once you're done with the program, just close the window.

