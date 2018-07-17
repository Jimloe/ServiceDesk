# ServiceDesk
Compilation of PowerShell functions within a GUI to make my job easier.

# ChangeLog
# v2.7.1 - 20180717
#-ComputerCleanup - Now works for the DS domain.  Matches any computers listed in the SchoolMatcher function and moves it to the proper OU.
#-Added the ability for the Logger function to automatically scroll as text is being displayed in the output box.
#-Added KTown, Landstuhl, Vogelweh and Ramstein schools to the supported list
#
# v2.7 - 20180717
#-MigrateComputers - Fix an error with the looping of the pingable computers resulting in a bad computername being passed to the next computer in the loop.
#
# v2.6 - 20180716
#-Added Brussels to the supported schools.
#-Tweaked the way the elevated.txt file is used/deleted to make it more seamless.
#-Tweaked SchoolMatcher to have it be same format in each new/old/prompt matcher.
#
# v2.5 - 20180713
#-Made some minor changes to the script layout for easier editing
#-Disabled some of the unfinished script function - DriveMappings, ComputerCleanup, BackupUserFiles
#-Modified the location of the ServiceDeskLogFile so it'll create the file under the root of the logged in user to prevent permission issues.
#-Modified the name of the file used to elevate the script to prevent confusion with the actual logging file.  New name is ElevateScript.txt, which is still located in the root of the current logged in user directory.
#-MigrateComputers - added the ability to exit out prematurely
#-SchoolMatcher - Added in a listing of the different drive mappings for the Schools.  Going to incorperate this into the DriveMappings function.
#-Logger - Added the ability to add a newline
#-DriveMappings - Finished the functionality for different sites to map their own drives.  Not all schools are complete. Need additional information.
#-Re-added logging to the main form.
#
# v2.5 - 20180713
#-Made some minor changes to the script layout for easier editing
#-Disabled some of the unfinished script function - DriveMappings, ComputerCleanup, BackupUserFiles
#-Modified the location of the ServiceDeskLogFile so it'll create the file under the root of the logged in user to prevent permission issues.
#-Modified the name of the file used to elevate the script to prevent confusion with the actual logging file.  New name is ElevateScript.txt, which is still located in the root of the current logged in user directory.
#-MigrateComputers - added the ability to exit out prematurely
#-SchoolMatcher - Added in a listing of the different drive mappings for the Schools.  Going to incorperate this into the DriveMappings function.
#-Logger - Added the ability to add a newline
#-DriveMappings - Finished the functionality for different sites to map their own drives.  Not all schools are complete. Need additional information.
#-Re-added logging to the main form.
#
# v2.4 - 20180712
#-Fixed an access denied error message.  The logger functionality in the Generate-Form was trying to run before the script was elevated causing access denied messages to users.
#-Added in the functionality to check to see if the user had already entered in EU & DS Credentials and won't prompt them again, saving some time when running MigrateComputers multiple times.
#-MigrateComputers - Fixed an error with having multiple rename computers.  The batch file wasn't ascii encoding causing problems with multiple lines.
#-MigrateComputers - Added the funcitonality to automatically select the right OU based on user location to move DS Joined computers to.
#-MigrateComputers - Slightly modified the FixWireless functionality to better log what the script is doing.
#
# v2.3 - 20180710
#-Added addtional functionality to the Logger function to streamline the script
#-Completed SchoolMatcher function to streamline script
#-MigrateComputers function uses SchoolMatcher function now
#-MigrateComputers function - Added the ability to select which domain you want to rename computers on
#-MigrateComputers function - Added the ability to select which domain you want to try and fix wireless settings on.  Need to test.
#-MigrateComputers function - Added a check to see if there were any computers to rename when running the batch file.
#-Completed GenerateCSV function to query ActiveDirectory for Win10 computers based on which Domain the user selects
#
# v2.2 - 20180705
#-Modified MigrateComputers so it only prompts for DS credentials once.  Additionally added Remove-ADComputer was added as a failsafe as the -UnjoinDomainCredential was not deleting the old computer object.
#
# v2.1 - 20180703
#-Implemented Logging capability and added it to functions
#
# v2.0 - 20180702
#Included Commands
#MigrateComputers, Enable-WinRM
#-Improved the GUI
#-Changed the RenameComputers and integrated that functionality into the MigrateComputers to cover everything with the DS Migration
#-MigrateComputers which can fix the wireless profile on laptops, migrate computers from the EU to the DS domain, and rename the computer to the proper naming convention.
#-Added About form and Helpful Links form
#
# v1.0 - 20180606
#Included Commands
#SCCM-Repair, Reset-MachinePW, CAC-Removal, GP-Update, ServerInfo, RenameCompters, DriveMappings (Only maps H: Drive at the moment)
#
