#Created and maintained by Jim Lower - james.lower.ctr@dodea.edu
#
#region ChangeLog
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
#Planned future functionality
#-Backup User files
#-Cleanup user profiles & registry keys & Hard Drive
#-User Drive mappings based on Group Membership
#-Add in the ability to pick which domain you want to work on
#endregion ChangeLog

#region Global Variables
$LocalSystemInfo = (Get-WmiObject -class Win32_OperatingSystem).Caption
$CurrentUser = Get-WMIObject -class Win32_ComputerSystem | select username
$CurrentUserName = $CurrentUser.username -replace ('^[\w]+\\','') 
$Matchme = $CurrentUser.username -match ('^[\w]+')
$CurrentUserDomain = $Matches.Values
$currentdate = Get-Date
$LogFileLoc = "C:\Users\$CurrentUserName\ServiceDeskLogFile.log"
$ErrorActionPreference = "Stop"
$CACPolicyRButton = ""
$EUCreds = $null
$DSCreds = $null
# Credits to - http://powershell.cz/2013/04/04/hide-and-show-console-window-from-gui/
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
#endregion Global Variables

#region Global Functions
function Logger{
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [string]$FilePath,
        [switch]$OPWindow,
        [switch]$NewLine
    )
    $timestamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
    $LogLine = "$timestamp| $Text"
    if($FilePath)
    {
        if(Test-Path $FilePath){Add-Content $FilePath -Value $LogLine}
        else
        {
        New-Item -Path $FilePath -ItemType File
        Add-Content $FilePath -Value $LogLine
    }
    }
    if($OPWindow.IsPresent)
    {
        $outputBox.AppendText("$Text")
        $outputBox.AppendText("`n")
    }
    if($NewLine.IsPresent)
    {$outputBox.AppendText("`n")}
}
function Show-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 5)
}
function Hide-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
}
function OpenFile {
    Logger -FilePath $LogFileLoc -Text "Prompting user to open .CSV"
	$selectOpenForm = New-Object System.Windows.Forms.OpenFileDialog
	$selectOpenForm.Filter = "CSV File (*.csv)|*.csv"
	$selectOpenForm.InitialDirectory = ".\"
	$selectOpenForm.Title = "Select a .CSV file which contains a list of computers to rename"
	$getKey = $selectOpenForm.ShowDialog()
	If ($getKey -eq "OK") {
            $inputFileName = $selectOpenForm.FileName
            $inputFileName
            Logger -FilePath $LogFileLoc -Text "OpenFile function outputting"
	}
    else {"0"}   
}
function directorysearcher{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Filter,
        [Parameter(Mandatory=$true)]
        [validateset('subtree','onelevel')]
        [string]$SearchScope,
        [Parameter(Mandatory=$true)]
        [validateset('FindOne','FindAll')]
        [string]$SearchType,
        [Parameter(Mandatory=$true)]
        [String]$SearchLocation
        )   
    [ADSI]$SearcherLoc = "LDAP://" + "$SearchLocation"
    $search = New-Object System.DirectoryServices.DirectorySearcher([ADSI]$SearcherLoc)
    $search.PageSize = 1000
    $search.Filter = $Filter
    $results = $search.$SearchType()
    $results
    Logger -FilePath $LogFileLoc -Text "directorysearcher function outputting"
}
function WaitProcess {  
     [CmdletBinding()]  
     param  
     (  
          [ValidateNotNullOrEmpty()][string]  
          $ProcessName  
     )  
     $Process = Get-Process $ProcessName -ErrorAction SilentlyContinue  
     $Process = $Process | Where-Object {$_.ProcessName -eq $ProcessName}  
     If($Process -ne $null){  
          Do{  
               Start-Sleep -Seconds 2  
               $Process = Get-Process $ProcessName -ErrorAction SilentlyContinue  
               $Process = $Process | Where-Object {$_.ProcessName -eq $ProcessName}  
          }  
          While ($Process -ne $null)  
     }
}
function SchoolMatcher{
    Param(
        [Parameter(Mandatory=$true)]
        [String]$CompName
    )
    $op = '' | Select OldName,NewName,Location,Room,User,Type,Barcode,OU,OldOU,MDrive,IDrive,HomeDrive
    #Old
    Logger -FilePath $LogFileLoc -Text "Attempting to match old naming convention" -OPWindow
    $matchme = $CompName -match "^(?<School>[\w]{4})-(?<Type>[\w]{1})(?<User>[\w]{1})(?<Room>[^.]{4})-(?<Barcode>[\d]{3})"
    if($matchme -eq $true)
    {
        Logger -FilePath $LogFileLoc -Text "Old naming convention matched" -OPWindow
        if($Matches.School -eq "ANES"){$newname = "ANSE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Ansbach ES"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Ansbach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Ansbach ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\ANSB-ANSB"  ; $IDrive = "\\ds.dodea.edu\Europe\ANSB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\ANSB-HOME" }#Ansbach ES
        if($Matches.School -eq "ANHS"){$newname = "ANSD" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Ansbach HS"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Ansbach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Ansbach MHS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"          ; $MDrive = "\\ds.dodea.edu\Europe\ANSB-ANSB"  ; $IDrive = "\\ds.dodea.edu\Europe\ANSB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\ANSB-HOME" }#Ansbach HS
        if($Matches.School -eq "AUES"){$newname = "AUKE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Wiesbaden ES"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Aukamm ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"            ; $MDrive = "\\ds.dodea.edu\Europe\WIES-WIMES" ; $IDrive = "\\ds.dodea.edu\Europe\WEIS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\WIES-HOME" }#Wiesbaden ES
        if($Matches.School -eq "BMHS"){$newname = "BAUD" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Baumholder HS"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Baumholder Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU" ; $OldOU = "OU=Windows10,OU=Computers,OU=Grafenwoehr ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"       ; $MDrive = "\\ds.dodea.edu\Europe\BAHR-BAHR"  ; $IDrive = "\\ds.dodea.edu\Europe\BAHR-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\BAHR-HOME" }#Baumholder HS
        if($Matches.School -eq "GAEM"){$newname = "GARD" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Garmisch MS"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Garmisch Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"   ; $OldOU = "OU=Windows10,OU=Computers,OU=Garmisch EMS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\GARM-GARM"  ; $IDrive = "" ; $HomeDrive = "\\ds.dodea.edu\Europe\GARM-HOME"                                     }#Garmisch MS                        
        if($Matches.School -eq "GRES"){$newname = "GRAE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Grafenwoehr ES" ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Grafenwoehr Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"; $OldOU = "OU=Windows10,OU=Computers,OU=Grafenwoehr ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"       ; $MDrive = "\\ds.dodea.edu\Europe\GRAF-GRAF"  ; $IDrive = "\\ds.dodea.edu\Europe\GRAF-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\GRAF-Home" }#Grafenwoehr ES
        if($Matches.School -eq "HOES"){$newname = "HOHE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Hoenfels ES"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Hohenfels Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Hohenfels ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\HOHE-HOHE"  ; $IDrive = "\\ds.dodea.edu\Europe\HOHE-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\HOHE-HOME" }#Hoenfels ES
        if($Matches.School -eq "HOHS"){$newname = "HOHD" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Hoenfels HS"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Hohenfels Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Hohenfels MHS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"        ; $MDrive = "\\ds.dodea.edu\Europe\HOHE-HOHE"  ; $IDrive = "\\ds.dodea.edu\Europe\HOHE-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\HOHE-HOME" }#Hoenfels HS
        if($Matches.School -eq "NZES"){$newname = "NETE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Netzaberg ES"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Grafenwoehr Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"; $OldOU = "OU=Windows10,OU=Computers,OU=Netzaberg ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\NETZ-NETZ"  ; $IDrive = "" ; $HomeDrive = "\\ds.dodea.edu\Europe\NETZ-HOME"                                     }#Netzaberg ES                       
        if($Matches.School -eq "NZMS"){$newname = "NETM" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Netzaberg MS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Grafenwoehr Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"; $OldOU = "OU=Windows10,OU=Computers,OU=Netzaberg MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\NETZ-NETZ"  ; $IDrive = "" ; $HomeDrive = "\\ds.dodea.edu\Europe\NETZ-HOME"                                     }#Netzaberg MS                       
        if($Matches.School -eq "PATC"){$newname = "PATM" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Patch MS"       ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Patch MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"             ; $MDrive = "\\ds.dodea.edu\Europe\STUT-PAMS"  ; $IDrive = "" ; $HomeDrive = ""                                                                    }#Patch MS                                                          
        if($Matches.School -eq "ROBI"){$newname = "ROBD" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Robinson ES"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Robinson Barracks ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU" ; $MDrive = "\\ds.dodea.edu\Europe\ROBI-ROBI"  ; $IDrive = "\\ds.dodea.edu\Europe\ROBI-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\ROBI-Home" }#Robinson ES
        if($Matches.School -eq "SBES"){$newname = "SEME" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Sembach ES"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Sembach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Sembach ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\SEMB-SEMB"  ; $IDrive = "\\ds.dodea.edu\Europe\SEMB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\SEMB-HOME" }#Sembach ES
        if($Matches.School -eq "SBMS"){$newname = "SEMM" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Sembach MS"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Sembach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Sembach MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\SEMB-SEMB"  ; $IDrive = "\\ds.dodea.edu\Europe\SEMB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\SEMB-HOME" }#Sembach MS
        if($Matches.School -eq "SMES"){$newname = "SMIE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Smith ES"       ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Baumholder Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU" ; $OldOU = "OU=Windows10,OU=Computers,OU=Smith ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"             ; $MDrive = "\\ds.dodea.edu\Europe\BAUM-SMIT"  ; $IDrive = "\\ds.dodea.edu\Europe\BAUM-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\BAUM-HOME" }#Smith ES
        if($Matches.School -eq "PAES"){$newname = "PATE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Patch ES"       ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Patch ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"             ; $MDrive = "" ; $IDrive = "" ; $HomeDrive = ""                                                                                                    }#Patch ES                                                                                         
        if($Matches.School -eq "STES"){$newname = "STUE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Stuttgart ES"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Stuttgart ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "" ; $IDrive = "\\ds.dodea.edu\Europe\STUT-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\STUT-Home"                                 }#Stuttgart ES                  
        if($Matches.School -eq "STHS"){$newname = "STUH" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Stuttgart HS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Stuttgart HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\STUT-STHS"  ; $IDrive = "\\ds.dodea.edu\Europe\STUT-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\STUT-Home" }#Stuttgart HS
        if($Matches.School -eq "VLES"){$newname = "VILE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Vilseck ES"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Vilseck Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Vilseck ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\VICE-VLES"  ; $IDrive = "\\ds.dodea.edu\Europe\VILS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\VILS-HOME" }#Vilseck ES
        if($Matches.School -eq "VLHS"){$newname = "VILD" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Vilseck HS"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Vilseck Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Vilseck HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\VICE-VICE"  ; $IDrive = "\\ds.dodea.edu\Europe\VILS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\VILS-HOME" }#Vilseck HS
        if($Matches.School -eq "HAES"){$newname = "HAIE" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Hainerberg ES"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Hainerberg ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"        ; $MDrive = "" ; $IDrive = "" ; $HomeDrive = ""                                                                                                    }#Hainerberg ES                                                                                    
        if($Matches.School -eq "WIHS"){$newname = "WIEH" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Wiesbaden HS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Wiesbaden HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\WIES-WIHS"  ; $IDrive = "\\ds.dodea.edu\Europe\WEIS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\WIES-HOME" }#Wiesbaden HS
        if($Matches.School -eq "WIMS"){$newname = "WIEM" + "-" + $Matches.Room + "-" + $Matches.Barcode + $Matches.Type + $Matches.User ; $SchoolLoc = "Wiesbaden MS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Wiesbaden MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\WIES-WIMES" ; $IDrive = "\\ds.dodea.edu\Europe\WEIS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\WIES-HOME" }#Wiesbaden MS
        $op.OldName   = $CompName
        $op.NewName   = $newname
        $op.Location  = $SchoolLoc
        $op.Room      = $RoomNum
        $op.User      = $Matches.User
        $op.Type      = $Matches.Type
        $op.Barcode   = $Matches.Barcode
        $op.OU        = $OU
        $op.OldOU     = $OldOU
        $op.MDrive    = $MDrive
        $op.IDrive    = $IDrive
        $op.HomeDrive = $HomeDrive
        $op
    }
    else{Logger -FilePath $LogFileLoc -Text "No match for old name" -OPWindow}
    #New
    if($matchme -eq $false)
    {
        Logger -FilePath $LogFileLoc -Text "Attempting to match new naming convention" -OPWindow
        $matchme = $CompName -match "^(?<School>[\w]{4})-(?<Room>[\w]{4})-(?<Barcode>[\d]{3})(?<Type>[\w]{1})(?<user>[\w]{1})"
        if($matchme -eq $true)
        {
            Logger -FilePath $LogFileLoc -Text "New naming convention matched" -OPWindow
            if($Matches.School -eq "ANSE"){$SchoolLoc = "Ansbach ES"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Ansbach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Ansbach ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\ANSB-ANSB"  ; $IDrive = "\\ds.dodea.edu\Europe\ANSB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\ANSB-HOME" }#Ansbach ES
            if($Matches.School -eq "ANSD"){$SchoolLoc = "Ansbach HS"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Ansbach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Ansbach MHS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"          ; $MDrive = "\\ds.dodea.edu\Europe\ANSB-ANSB"  ; $IDrive = "\\ds.dodea.edu\Europe\ANSB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\ANSB-HOME" }#Ansbach HS
            if($Matches.School -eq "AUKE"){$SchoolLoc = "Wiesbaden ES"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Aukamm ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"            ; $MDrive = "\\ds.dodea.edu\Europe\WIES-WIMES" ; $IDrive = "\\ds.dodea.edu\Europe\WEIS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\WIES-HOME" }#Wiesbaden ES
            if($Matches.School -eq "BAUD"){$SchoolLoc = "Baumholder HS"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Baumholder Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU" ; $OldOU = "OU=Windows10,OU=Computers,OU=Grafenwoehr ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"       ; $MDrive = "\\ds.dodea.edu\Europe\BAHR-BAHR"  ; $IDrive = "\\ds.dodea.edu\Europe\BAHR-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\BAHR-HOME" }#Baumholder HS
            if($Matches.School -eq "GARD"){$SchoolLoc = "Garmisch MS"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Garmisch Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"   ; $OldOU = "OU=Windows10,OU=Computers,OU=Garmisch EMS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\GARM-GARM"  ; $IDrive = "" ; $HomeDrive = "\\ds.dodea.edu\Europe\GARM-HOME"                                     }#Garmisch MS   
            if($Matches.School -eq "GRAE"){$SchoolLoc = "Grafenwoehr ES" ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Grafenwoehr Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"; $OldOU = "OU=Windows10,OU=Computers,OU=Grafenwoehr ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"       ; $MDrive = "\\ds.dodea.edu\Europe\GRAF-GRAF"  ; $IDrive = "\\ds.dodea.edu\Europe\GRAF-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\GRAF-Home" }#Grafenwoehr ES
            if($Matches.School -eq "HOHE"){$SchoolLoc = "Hoenfels ES"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Hohenfels Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Hohenfels ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\HOHE-HOHE"  ; $IDrive = "\\ds.dodea.edu\Europe\HOHE-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\HOHE-HOME" }#Hoenfels ES
            if($Matches.School -eq "HOHD"){$SchoolLoc = "Hoenfels HS"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Hohenfels Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Hohenfels MHS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"        ; $MDrive = "\\ds.dodea.edu\Europe\HOHE-HOHE"  ; $IDrive = "\\ds.dodea.edu\Europe\HOHE-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\HOHE-HOME" }#Hoenfels HS
            if($Matches.School -eq "NETE"){$SchoolLoc = "Netzaberg ES"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Grafenwoehr Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"; $OldOU = "OU=Windows10,OU=Computers,OU=Netzaberg ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\NETZ-NETZ"  ; $IDrive = "" ; $HomeDrive = "\\ds.dodea.edu\Europe\NETZ-HOME"                                     }#Netzaberg ES  
            if($Matches.School -eq "NETM"){$SchoolLoc = "Netzaberg MS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Grafenwoehr Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"; $OldOU = "OU=Windows10,OU=Computers,OU=Netzaberg MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\NETZ-NETZ"  ; $IDrive = "" ; $HomeDrive = "\\ds.dodea.edu\Europe\NETZ-HOME"                                     }#Netzaberg MS  
            if($Matches.School -eq "PATM"){$SchoolLoc = "Patch MS"       ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Patch MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"             ; $MDrive = "\\ds.dodea.edu\Europe\STUT-PAMS"  ; $IDrive = "" ; $HomeDrive = ""                                                                    }#Patch MS      
            if($Matches.School -eq "ROBD"){$SchoolLoc = "Robinson ES"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Robinson Barracks ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU" ; $MDrive = "\\ds.dodea.edu\Europe\ROBI-ROBI"  ; $IDrive = "\\ds.dodea.edu\Europe\ROBI-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\ROBI-Home" }#Robinson ES
            if($Matches.School -eq "SEME"){$SchoolLoc = "Sembach ES"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Sembach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Sembach ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\SEMB-SEMB"  ; $IDrive = "\\ds.dodea.edu\Europe\SEMB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\SEMB-HOME" }#Sembach ES
            if($Matches.School -eq "SEMM"){$SchoolLoc = "Sembach MS"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Sembach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Sembach MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\SEMB-SEMB"  ; $IDrive = "\\ds.dodea.edu\Europe\SEMB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\SEMB-HOME" }#Sembach MS
            if($Matches.School -eq "SMIE"){$SchoolLoc = "Smith ES"       ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Baumholder Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU" ; $OldOU = "OU=Windows10,OU=Computers,OU=Smith ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"             ; $MDrive = "\\ds.dodea.edu\Europe\BAUM-SMIT"  ; $IDrive = "\\ds.dodea.edu\Europe\BAUM-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\BAUM-HOME" }#Smith ES
            if($Matches.School -eq "PATE"){$SchoolLoc = "Patch ES"       ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Patch ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"             ; $MDrive = "" ; $IDrive = "" ; $HomeDrive = ""                                                                                                    }#Patch ES      
            if($Matches.School -eq "STUE"){$SchoolLoc = "Stuttgart ES"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Stuttgart ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "" ; $IDrive = "\\ds.dodea.edu\Europe\STUT-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\STUT-Home"                                 }#Stuttgart ES  
            if($Matches.School -eq "STUH"){$SchoolLoc = "Stuttgart HS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Stuttgart HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\STUT-STHS"  ; $IDrive = "\\ds.dodea.edu\Europe\STUT-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\STUT-Home" }#Stuttgart HS
            if($Matches.School -eq "VILE"){$SchoolLoc = "Vilseck ES"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Vilseck Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Vilseck ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\VICE-VLES"  ; $IDrive = "\\ds.dodea.edu\Europe\VILS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\VILS-HOME" }#Vilseck ES
            if($Matches.School -eq "VILD"){$SchoolLoc = "Vilseck HS"     ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Vilseck Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Vilseck HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\VICE-VICE"  ; $IDrive = "\\ds.dodea.edu\Europe\VILS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\VILS-HOME" }#Vilseck HS
            if($Matches.School -eq "HAIE"){$SchoolLoc = "Hainerberg ES"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Hainerberg ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"        ; $MDrive = "" ; $IDrive = "" ; $HomeDrive = ""                                                                                                    }#Hainerberg ES 
            if($Matches.School -eq "WIEH"){$SchoolLoc = "Wiesbaden HS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Wiesbaden HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\WIES-WIHS"  ; $IDrive = "\\ds.dodea.edu\Europe\WEIS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\WIES-HOME" }#Wiesbaden HS
            if($Matches.School -eq "WIEM"){$SchoolLoc = "Wiesbaden MS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Wiesbaden MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\WIES-WIMES" ; $IDrive = "\\ds.dodea.edu\Europe\WEIS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\WIES-HOME" }#Wiesbaden MS
            $op.OldName  = $CompName
            $op.NewName  = ""
            $op.Location = $SchoolLoc
            $op.Room     = $RoomNum
            $op.User     = $Matches.User
            $op.Type     = $Matches.Type
            $op.Barcode  = $Matches.Barcode
            $op.OU       = $OU
            $op.OldOU    = $OldOU
            $op.MDrive    = $MDrive
            $op.IDrive    = $IDrive
            $op.HomeDrive = $HomeDrive
            $op
        }
        #Prompt User for new computer name as we're unable to match the old one
        else
        {
            $Form.TopMost = $false
            Logger -FilePath $LogFileLoc -Text "Unable to match computer name.  Prompting user for new computer name." -OPWindow
            do {
                [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                $ComputerInput = [Microsoft.VisualBasic.Interaction]::InputBox("Unable to match $CompName, please enter in what you want the computer name to be in the new valid naming format.", "Bad Computer Name","")
                $matchme = $ComputerInput -match "^(?<School>[\w]{4})-(?<Room>[\w]{4})-(?<Barcode>[\d]{3})(?<Type>[\w]{1})(?<user>[\w]{1})"

            } until ($matchme -eq $true)
            Logger -FilePath $LogFileLoc -Text "User has entered in $ComputerInput as new computer name" -OPWindow
            Logger -FilePath $LogFileLoc -Text "Matching new computer name" -OPWindow
            $Matches
            if($Matches.School -eq "ANSE"){$SchoolLoc = "Ansbach ES"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Ansbach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Ansbach ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\ANSB-ANSB"  ; $IDrive = "\\ds.dodea.edu\Europe\ANSB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\ANSB-HOME" }#Ansbach ES
            if($Matches.School -eq "ANSD"){$SchoolLoc = "Ansbach HS"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Ansbach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Ansbach MHS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"          ; $MDrive = "\\ds.dodea.edu\Europe\ANSB-ANSB"  ; $IDrive = "\\ds.dodea.edu\Europe\ANSB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\ANSB-HOME" }#Ansbach HS
            if($Matches.School -eq "AUKE"){$SchoolLoc = "Wiesbaden ES"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Aukamm ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"            ; $MDrive = "\\ds.dodea.edu\Europe\WIES-WIMES" ; $IDrive = "\\ds.dodea.edu\Europe\WEIS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\WIES-HOME" }#Wiesbaden ES
            if($Matches.School -eq "BAUD"){$SchoolLoc = "Baumholder HS" ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Baumholder Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU" ; $OldOU = "OU=Windows10,OU=Computers,OU=Grafenwoehr ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"       ; $MDrive = "\\ds.dodea.edu\Europe\BAHR-BAHR"  ; $IDrive = "\\ds.dodea.edu\Europe\BAHR-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\BAHR-HOME" }#Baumholder HS
            if($Matches.School -eq "GARD"){$SchoolLoc = "Garmisch MS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Garmisch Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"   ; $OldOU = "OU=Windows10,OU=Computers,OU=Garmisch EMS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\GARM-GARM"  ; $IDrive = "" ; $HomeDrive = "\\ds.dodea.edu\Europe\GARM-HOME"                                     }#Garmisch MS   
            if($Matches.School -eq "GRAE"){$SchoolLoc = "Grafenwoehr ES"; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Grafenwoehr Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"; $OldOU = "OU=Windows10,OU=Computers,OU=Grafenwoehr ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"       ; $MDrive = "\\ds.dodea.edu\Europe\GRAF-GRAF"  ; $IDrive = "\\ds.dodea.edu\Europe\GRAF-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\GRAF-Home" }#Grafenwoehr ES
            if($Matches.School -eq "HOHE"){$SchoolLoc = "Hoenfels ES"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Hohenfels Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Hohenfels ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\HOHE-HOHE"  ; $IDrive = "\\ds.dodea.edu\Europe\HOHE-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\HOHE-HOME" }#Hoenfels ES
            if($Matches.School -eq "HOHD"){$SchoolLoc = "Hoenfels HS"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Hohenfels Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Hohenfels MHS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"        ; $MDrive = "\\ds.dodea.edu\Europe\HOHE-HOHE"  ; $IDrive = "\\ds.dodea.edu\Europe\HOHE-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\HOHE-HOME" }#Hoenfels HS
            if($Matches.School -eq "NETE"){$SchoolLoc = "Netzaberg ES"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Grafenwoehr Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"; $OldOU = "OU=Windows10,OU=Computers,OU=Netzaberg ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\NETZ-NETZ"  ; $IDrive = "" ; $HomeDrive = "\\ds.dodea.edu\Europe\NETZ-HOME"                                     }#Netzaberg ES  
            if($Matches.School -eq "NETM"){$SchoolLoc = "Netzaberg MS"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Grafenwoehr Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"; $OldOU = "OU=Windows10,OU=Computers,OU=Netzaberg MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\NETZ-NETZ"  ; $IDrive = "" ; $HomeDrive = "\\ds.dodea.edu\Europe\NETZ-HOME"                                     }#Netzaberg MS  
            if($Matches.School -eq "PATM"){$SchoolLoc = "Patch MS"      ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Patch MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"             ; $MDrive = "\\ds.dodea.edu\Europe\STUT-PAMS"  ; $IDrive = "" ; $HomeDrive = ""                                                                    }#Patch MS      
            if($Matches.School -eq "ROBD"){$SchoolLoc = "Robinson ES"   ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Robinson Barracks ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU" ; $MDrive = "\\ds.dodea.edu\Europe\ROBI-ROBI"  ; $IDrive = "\\ds.dodea.edu\Europe\ROBI-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\ROBI-Home" }#Robinson ES
            if($Matches.School -eq "SEME"){$SchoolLoc = "Sembach ES"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Sembach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Sembach ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\SEMB-SEMB"  ; $IDrive = "\\ds.dodea.edu\Europe\SEMB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\SEMB-HOME" }#Sembach ES
            if($Matches.School -eq "SEMM"){$SchoolLoc = "Sembach MS"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Sembach Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Sembach MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\SEMB-SEMB"  ; $IDrive = "\\ds.dodea.edu\Europe\SEMB-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\SEMB-HOME" }#Sembach MS
            if($Matches.School -eq "SMIE"){$SchoolLoc = "Smith ES"      ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Baumholder Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU" ; $OldOU = "OU=Windows10,OU=Computers,OU=Smith ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"             ; $MDrive = "\\ds.dodea.edu\Europe\BAUM-SMIT"  ; $IDrive = "\\ds.dodea.edu\Europe\BAUM-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\BAUM-HOME" }#Smith ES
            if($Matches.School -eq "PATE"){$SchoolLoc = "Patch ES"      ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Patch ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"             ; $MDrive = "" ; $IDrive = "" ; $HomeDrive = ""                                                                                                    }#Patch ES      
            if($Matches.School -eq "STUE"){$SchoolLoc = "Stuttgart ES"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Stuttgart ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "" ; $IDrive = "\\ds.dodea.edu\Europe\STUT-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\STUT-Home"                                 }#Stuttgart ES  
            if($Matches.School -eq "STUH"){$SchoolLoc = "Stuttgart HS"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Stuttgart Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Stuttgart HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\STUT-STHS"  ; $IDrive = "\\ds.dodea.edu\Europe\STUT-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\STUT-Home" }#Stuttgart HS
            if($Matches.School -eq "VILE"){$SchoolLoc = "Vilseck ES"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Vilseck Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Vilseck ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\VICE-VLES"  ; $IDrive = "\\ds.dodea.edu\Europe\VILS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\VILS-HOME" }#Vilseck ES
            if($Matches.School -eq "VILD"){$SchoolLoc = "Vilseck HS"    ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Vilseck Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"    ; $OldOU = "OU=Windows10,OU=Computers,OU=Vilseck HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"           ; $MDrive = "\\ds.dodea.edu\Europe\VICE-VICE"  ; $IDrive = "\\ds.dodea.edu\Europe\VILS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\VILS-HOME" }#Vilseck HS
            if($Matches.School -eq "HAIE"){$SchoolLoc = "Hainerberg ES" ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Hainerberg ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"        ; $MDrive = "" ; $IDrive = "" ; $HomeDrive = ""                                                                                                    }#Hainerberg ES 
            if($Matches.School -eq "WIEH"){$SchoolLoc = "Wiesbaden HS"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Wiesbaden HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\WIES-WIHS"  ; $IDrive = "\\ds.dodea.edu\Europe\WEIS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\WIES-HOME" }#Wiesbaden HS
            if($Matches.School -eq "WIEM"){$SchoolLoc = "Wiesbaden MS"  ; $RoomNum = $Matches.Room ; $OU = "OU=Production,OU=Computers,OU=Wiesbaden Germany,OU=East,OU=Europe,OU=Regional Resources,DC=DS,DC=DODEA,DC=EDU"  ; $OldOU = "OU=Windows10,OU=Computers,OU=Wiesbaden MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU"         ; $MDrive = "\\ds.dodea.edu\Europe\WIES-WIMES" ; $IDrive = "\\ds.dodea.edu\Europe\WEIS-InfoDrive" ; $HomeDrive = "\\ds.dodea.edu\Europe\WIES-HOME" }#Wiesbaden MS
            $op.OldName  = $CompName
            $op.NewName  = $ComputerInput
            $op.Location = $SchoolLoc
            $op.Room     = $RoomNum
            $op.User     = $Matches.User
            $op.Type     = $Matches.Type
            $op.Barcode  = $Matches.Barcode
            $op.OU       = $OU
            $op.OldOU    = $OldOU
            $op.MDrive    = $MDrive
            $op.IDrive    = $IDrive
            $op.HomeDrive = $HomeDrive
            $op
            $Form.TopMost = $true
        }
    }
    Logger -FilePath $LogFileLoc -Text "SchoolMatcher function exiting with $op"
}
function Enable-WinRM{
    [CmdletBinding(SupportsShouldProcess = $true,
                   PositionalBinding = $true,
                   ConfirmImpact = 'Medium')]
    [OutputType([Boolean])]
    Param
    (
        # The computer to process
        [Parameter(Mandatory = $true,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName = $true,
                   ValueFromRemainingArguments = $false,
                   Position = 0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("cn,computername,pscomputername,server")]
        [string]
        $ComputerName,
        # The number of seconds to wait before timing out
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName = $true,
                   ValueFromRemainingArguments = $false,
                   Position = 1)]
        [AllowNull()]
        [AllowEmptyString()]
        [int]
        $Timeout = 30
    )
    Begin {
        if (-not (Test-Connection $ComputerName -Quiet)) {
            Write-Error "Unable to connect to $ComputerName"
            return $false
        }
    }
    Process {
        if ($pscmdlet.ShouldProcess("$ComputerName", "Enable WinRM")) {
            try {
                if (Test-WSMan $ComputerName -ErrorAction Stop) {
                    Write-Verbose "WSMan already enabled on $ComputerName"
                    return $true
                }
            }
            catch {
                Invoke-WmiMethod -class win32_Process -name create -ArgumentList "C:\WINDOWS\system32\winrm.cmd quickconfig -q" -ComputerName $ComputerName | Out-Null
                $end = [DateTime]::Now.AddSeconds($Timeout)
                do {
                    Write-Verbose "Testing WSMan connection"
                    Start-Sleep -Seconds 5
                    if ([DateTime]::Now -gt $end) {
                        Write-Error "Timeout waiting for WinRM to enable on $ComputerName"
                        return $false
                    }
                }
                while (-not (Test-WSMan $ComputerName -ErrorAction SilentlyContinue))
                Write-Verbose "WSMan enabled on $ComputerName"
                return $true
            }
        }
    }
}
#endregion Global Functions

#region Work Functions
#These functions may reference Global Functions
function MigrateComputers{
    Param(
        [switch]$RenameComputers,
        [switch]$ChangeDomain,
        [switch]$FixWireless
    )
    #Stuff that every sub function needs
    #
    #Securely getting users password, and logged on user
    #####################################################################
    $outputBox.Text = ''
    if ($EUCreds -eq $null)
    {
        Logger -FilePath $LogFileLoc -Text "Prompting user for .SDT Credentials" -OPWindow
        $Form.TopMost = $false
        $PWIN = Get-Credential -UserName "eu.ds.dodea.edu\" -Message "Please input elevated EU credentials"
        if($PWIN -eq $null){Logger -FilePath $LogFileLoc -Text "User canceled out of window." -OPWindow ; Return}
    }
    if ($DSCreds -eq $null)
    {
        Logger -FilePath $LogFileLoc -Text "Prompting user for .CSS Credentials" -OPWindow
        $DSCreds = Get-Credential -UserName "ds.dodea.edu\" -Message "Please input elevated DS credentials"
        if($DSCreds -eq $null){Logger -FilePath $LogFileLoc -Text "User canceled out of window." -OPWindow ; Return}
    }
    Logger -FilePath $LogFileLoc -Text "Prompting user for domain" -OPWindow
    $matchme = $null
    do {
        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
        $DomainInput = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter in EU or DS to choose which domain you want to work on.", "Choose a domain","")
        $matchme = $DomainInput -match "EU|DS"
    
    } until ($matchme -eq $true)
    #Importing .CSV file
    #####################################################################
    Logger -FilePath $LogFileLoc -Text "Prompting user for .CSV input. Calling OpenFile function" -OPWindow
    $filelocationOpenFile = OpenFile
    if($filelocationOpenFile -eq "0"){Logger -FilePath $LogFileLoc -Text "Canceled by user." ; $Form.TopMost = $true ; Return}
    $computerlist = Import-Csv -Path $filelocationOpenFile -Header "Computers"
    #Looping through computers in the .CSV file
    #####################################################################
    Logger -FilePath $LogFileLoc -Text "Proccessing imported computer list" -OPWindow
    Logger -FilePath $LogFileLoc -Text "Creating jobs to ping multiple computers and creating Array for future use" -OPWindow
    $PingableItems = New-Object System.Collections.Generic.List[System.Object]
    Foreach($targetcomputer in $computerlist.Computers){Set-Variable -Name "Status_$targetcomputer" -Value (Test-Connection -ComputerName $targetcomputer -AsJob -Count 1)}
    Get-Variable "Status_*" -ValueOnly | Foreach {
        $Status = Wait-Job $_ | Receive-Job
        if ($Status.ResponseTime -ne $null ) {
            $PingableItems.Add($Status.Address)
        }
    }
    $PingableItems.ToArray()
    #Logger -FilePath $LogFileLoc -Text "Attempting to match hostnames and generate new ones" -OPWindow
    if($RenameComputers.IsPresent)
    {
        Logger -FilePath $LogFileLoc -Text "Generating Batch file for computer renaming." -OPWindow
        New-Item -Path C:\Users\$CurrentUserName -Name RenameRemoteComputers.bat -ItemType File -Force | Out-Null
    }
    #Checking pingable status
    #####################################################################
    Foreach($targetcomp in $PingableItems){
        $status = Test-Connection -ComputerName $targetcomp -Count 1 -ErrorAction Continue
        if($Status.ResponseTime -ne $null){
            Logger -FilePath $LogFileLoc -Text "Was able to ping $targetcomp, running SchoolMatch function" -OPWindow
            $SchoolMatcher = SchoolMatcher -CompName $targetcomputer
            #Doing work on computers here
            #####################################################################
            if($RenameComputers.IsPresent)
            {
                Logger -FilePath $LogFileLoc -Text "Generating command to rename computer" -OPWindow
                $command = "netdom.exe"
                $renamecomp = "renamecomputer"
                $compname = "/newname:$($SchoolMatcher.newname)"
                if($Matches.Values -eq "EU")
                {
                    $PWOUT = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($PWIN.Password))
                    $user = "/userd:$($PWIN.UserName)"
                    $pw = "/passwordd:$PWOUT" 
                }
                if($Matches.Values -eq "DS")
                {
                    $PWOUT = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($DSCreds.Password))
                    $user = "/userd:$($DSCreds.UserName)"
                    $pw = "/passwordd:$PWOUT" 
                }
                $reboot = "/reboot:1"
                $forceit = "/Force"
                $batgenerator = $command + ' ' + $renamecomp + ' ' + $targetcomp + ' ' + $compname + ' ' + $user + ' ' + $pw + ' ' + $reboot + ' ' + $forceit
                $batgenerator | Out-File -FilePath C:\Users\$CurrentUserName\RenameRemoteComputers.bat -Append -Encoding ascii
                Logger -FilePath $LogFileLoc -Text "Created Batch File with command"
                Logger -FilePath $LogFileLoc -Text "$targetcomp is going to be renamed to $newname" -OPWindow
            }
            if($ChangeDomain.IsPresent)
            {
                $domain  = 'ds.dodea.edu'
                $EUCreds = $PWIN
                Logger -FilePath $LogFileLoc -Text "Attempting to join $targetcomp to the DS Domain" -OPWindow
                $outputBox.AppendText("Attempting to join $targetcomp to the DS Domain using the new name of $($SchoolMatcher.newname)")
                try{Add-Computer -ComputerName $targetcomp -DomainName $domain -OUPath "$($SchoolMatcher.OU)" -Credential $DSCreds -NewName "$($SchoolMatcher.newname)" -UnjoinDomainCredential $EUCreds -Restart}
                catch{Logger -FilePath $LogFileLoc -Text "Failed with Error Message: $($Error[0])" -OPWindow -NewLine}
                try{
                    Remove-ADComputer -Identity $targetcomp -Credential $EUCreds
                    Logger -FilePath $LogFileLoc -Text "Removal succeeded" -OPWindow
                }catch{Logger -FilePath $LogFileLoc -Text "Removal failed with error message: $($Error[0])" -OPWindow -NewLine}
                
            }
            if($FixWireless.IsPresent)
            {
                try
                {
                    Logger -FilePath $LogFileLoc -Text "Attempting to enable WinRM on $targetcomp" -OPWindow
                    Enable-WinRM -ComputerName $targetcomp
                }
                catch{Logger -FilePath $LogFileLoc -Text "Failed Settings WinRM on $targetcomp with error message $($error[0])" -OPWindow -NewLine}
                try{
                    Logger -FilePath $LogFileLoc -Text "Successfully set WinRM Settings" -OPWindow
                    if($Matches.Values -eq "EU"){$wificreds = $PWIN}
                    if($Matches.Values -eq "DS"){$wificreds = $DSCreds}
                    Logger -FilePath $LogFileLoc -Text "Attempting to set wireless profile on $targetcomp" -OPWindow
                    $WifiResult = Invoke-Command -ComputerName $targetcomp -Credential $wificreds {
                        $WifiXML = @'
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
	<name>DSDoDEAWireless</name>
	<SSIDConfig>
		<SSID>
			<hex>446F444541576972656C657373</hex>
			<name>DoDEAWireless</name>
		</SSID>
		<nonBroadcast>true</nonBroadcast>
	</SSIDConfig>
	<connectionType>ESS</connectionType>
	<connectionMode>auto</connectionMode>
	<autoSwitch>false</autoSwitch>
	<MSM>
		<security>
			<authEncryption>
				<authentication>WPA2</authentication>
				<encryption>AES</encryption>
				<useOneX>true</useOneX>
			</authEncryption>
			<OneX xmlns="http://www.microsoft.com/networking/OneX/v1">
				<maxAuthFailures>1</maxAuthFailures>
				<authMode>machine</authMode>
				<EAPConfig><EapHostConfig xmlns="http://www.microsoft.com/provisioning/EapHostConfig"><EapMethod><Type xmlns="http://www.microsoft.com/provisioning/EapCommon">25</Type><VendorId xmlns="http://www.microsoft.com/provisioning/EapCommon">0</VendorId><VendorType xmlns="http://www.microsoft.com/provisioning/EapCommon">0</VendorType><AuthorId xmlns="http://www.microsoft.com/provisioning/EapCommon">0</AuthorId></EapMethod><Config xmlns="http://www.microsoft.com/provisioning/EapHostConfig"><Eap xmlns="http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"><Type>25</Type><EapType xmlns="http://www.microsoft.com/provisioning/MsPeapConnectionPropertiesV1"><ServerValidation><DisableUserPromptForServerValidation>true</DisableUserPromptForServerValidation><ServerNames></ServerNames></ServerValidation><FastReconnect>true</FastReconnect><InnerEapOptional>false</InnerEapOptional><Eap xmlns="http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1"><Type>26</Type><EapType xmlns="http://www.microsoft.com/provisioning/MsChapV2ConnectionPropertiesV1"><UseWinLogonCredentials>true</UseWinLogonCredentials></EapType></Eap><EnableQuarantineChecks>false</EnableQuarantineChecks><RequireCryptoBinding>false</RequireCryptoBinding><PeapExtensions><PerformServerValidation xmlns="http://www.microsoft.com/provisioning/MsPeapConnectionPropertiesV2">false</PerformServerValidation><AcceptServerName xmlns="http://www.microsoft.com/provisioning/MsPeapConnectionPropertiesV2">false</AcceptServerName></PeapExtensions></EapType></Eap></Config></EapHostConfig></EAPConfig>
			</OneX>
		</security>
	</MSM>
</WLANProfile>

'@
                        Set-Content -Value $WifiXML -Path C:\DSDodeaWireless.xml
                        netsh wlan add profile filename=c:\dsdodeawireless.xml interface="Wi-Fi"
                        netsh wlan set profileorder name="DSDoDEAWireless" interface="Wi-Fi" priority=1
                    }
                    Logger -FilePath $LogFileLoc -Text "$WifiResult" -OPWindow
                }
                Catch{Logger -FilePath $LogFileLoc -Text "Unable to set wifi profile on $targetcomp with error message $($Error[0])" -OPWindow -NewLine}
            }          
        }
        else{
            Logger -FilePath $LogFileLoc -Text "$targetcomp was not pingable, skipping it" -OPWindow
        }
    }
    #Run the batch file
    #####################################################################
    if($RenameComputers.IsPresent)
    {
        Logger -FilePath $LogFileLoc -Text "Running batch file to rename computers" -OPWindow
        if ((Get-Content -Path "C:\Users\$CurrentUserName\RenameRemoteComputers.bat") -ne $null)
        {
            try{
                Invoke-Expression "C:\Users\$CurrentUserName\RenameRemoteComputers.bat"
                Remove-Item "C:\Users\$CurrentUserName\RenameRemoteComputers.bat" -Force
                Logger -FilePath $LogFileLoc -Text "Finished running Batch file" -OPWindow
            }
            catch{Logger -FilePath $LogFileLoc -Text "Failed running the batch file with error message: $($Error[0])" -OPWindow -NewLine}
        }
        else{Logger -FilePath $LogFileLoc -Text "There were no computers on to rename." -OPWindow; Return}
    }
    Logger -FilePath $LogFileLoc -Text "Finished running MigrateComputers function"
    Logger -Text "Completed Migrating computers" -OPWindow
    $Form.TopMost = $true
}
function SCCM-Repair{
    Param(
        [switch]$Reinstall,
        [switch]$Site,
        [switch]$RunActions
        )
    if($Reinstall.IsPresent){
        $outputBox.Text = ''
        #List out the path and the management point arguments
        $runme = "$env:windir\ccmsetup\ccmsetup.exe"
        $UNargs = '/uninstall' 
        $INargs = '/mp:EU-SCCMP01.ds.dodea.edu'
        #Start the uninstall process
        Logger -FilePath $LogFileLoc -Text "Starting uninstall of CCM Client..." -OPWindow
        If ((Test-Path $runme) -eq $true) {
            $ErrCode = (Start-Process -FilePath $runme -ArgumentList $UNargs -WindowStyle Minimized -Wait -Passthru).ExitCode
            If (($ErrCode -eq 0) -or ($ErrCode -eq 3010))
            {  
                Logger -FilePath $LogFileLoc -Text "Uninstall of CCM Client complete" -OPWindow
                If ((Test-Path $runme) -eq $true) 
                {  
                    WaitProcess -ProcessName ccmsetup  
                } 
                else
                {
                    Logger -FilePath $LogFileLoc -Text "Uninstall of CCM Client failed" -OPWindow
                    $Failed = $true  
                }  
            }
            else
            {
                Logger -FilePath $LogFileLoc -Text "Failed with error, $ErrCode" -OPWindow
            } 
        }
        else
        {
            Logger -FilePath $LogFileLoc -Text "SCCM exectuable not present" -OPWindow
        }
    
        #Start the install process
        $ErrCode = (Start-Process -FilePath $runme -ArgumentList $INargs -WindowStyle Minimized -Wait -Passthru).ExitCode
        Logger -FilePath $LogFileLoc -Text "Starting install of CCM Client..." -OPWindow
        #Checks the return code of the software install
        If (($ErrCode -eq 0) -or ($ErrCode -eq 3010))
        {  
            $outputBox.AppendText("----Success")
            Logger -FilePath $LogFileLoc -Text "Success" -OPWindow
            If ((Test-Path $runme) -eq $true)
            {  
                WaitProcess -ProcessName ccmsetup  
            } 
            else
            {  
                Logger -FilePath $LogFileLoc -Text "Failed" -OPWindow
                $Failed = $true  
            }  
        }
        else
        {  
            Logger -FilePath $LogFileLoc -Text "Failed with error, $ErrCode" -OPWindow
        } 
        #Assigns the site code after the client has been installed
        SCCM-Repair -Site
        Start-Sleep -Seconds 2
        #Runs the Machine Policy Retrieval to check in with the SCCM server.
        Logger -FilePath $LogFileLoc -Text "Running the Machine Policy Retrieval..." -OPWindow
        try{
            Invoke-WmiMethod -Namespace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"
            Logger -FilePath $LogFileLoc -Text "Success" -OPWindow
            Logger -FilePath $LogFileLoc -Text "Reinstall of SCCM completed." -OPWindow
            } #Starts the Machine Policy Retrieval.  If it fails, it will retry in 10 seconds
        catch{
            Logger -FilePath $LogFileLoc -Text "Failed, trying again in 10 seconds" -OPWindow
            Start-Sleep -Seconds 10
            try{
                Invoke-WmiMethod -Namespace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"
                Logger -FilePath $LogFileLoc -Text "Success" -OPWindow
            }
            catch{
                Logger -FilePath $LogFileLoc -Text "Failed, try running the actions manually and ensure that the Computer has been approved on the SCCM Server" -OPWindow
            }
        }
    
    }
#####################################################################
#####Assigns the site code
    if($Site.IsPresent){
        $outputBox.Text = ''
        Logger -FilePath $LogFileLoc -Text "Checking SCCM Site code..." -OPWindow
        #Attempts to get the site code
        $sms = new-object –comobject “Microsoft.SMS.Client”
        Try{$SiteCode = $sms.GetAssignedSite()}
        Catch{
            Logger -FilePath $LogFileLoc -Text "No site found" -OPWindow
            }
        if($SiteCode -ne 'E01'){
            Logger -FilePath $LogFileLoc -Text "Assigning site code..E01" -OPWindow
            #
            #
            #
            #
            #
            #
            #
            #
            Logger -FilePath $LogFileLoc -Text "Success" -OPWindow
            }  #Checks to see if the site code is set to E01 and if not, assigns it to that     
        else{
            Logger -FilePath $LogFileLoc -Text "Already set to E01" -OPWindow
            }
        Logger -FilePath $LogFileLoc -Text "Checking SCCM Site code completed." -OPWindow
    }
#####Runs the Machine Policy Retrieval to check in with the SCCM server.
    if($RunActions.IsPresent){
        $outputBox.Text = ''
        Logger -FilePath $LogFileLoc -Text "Running the Machine Policy Retrieval..." -OPWindow
        try
        {
            Invoke-WmiMethod -Namespace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}" | Out-Null
            Logger -FilePath $LogFileLoc -Text "Success" -OPWindow
        } #Starts the Machine Policy Retrieval
        catch
        {
            Logger -FilePath $LogFileLoc -Text "Failed, trying again in 10 seconds" -OPWindow
            Start-Sleep -Seconds 10
            try
            {
                Invoke-WmiMethod -Namespace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}" | Out-Null
                Logger -FilePath $LogFileLoc -Text "Success" -OPWindow
            }
            catch
            {
                Logger -FilePath $LogFileLoc -Text "Failed, try running the actions manually and ensure that the Computer has been approved on the SCCM Server" -OPWindow
            }
        }
        Logger -FilePath $LogFileLoc -Text "Running client actions completed." -OPWindow
    }
#####################################################################
}
function Reset-MachinePW{
    $outputBox.Text = ""
    Logger -FilePath $LogFileLoc -Text "Grabbing Credentials." -OPWindow
    $creds = Get-Credential -Message "Enter .SDT Account info" -UserName "eu.ds.dodea.edu\$env:USERNAME"
    Logger -FilePath $LogFileLoc -Text "Querying local Domain Controller." -OPWindow
    $command = "nltest"
    $args = "/DSGETDC:"
    $QueryDC = & $command $args 
    $OutDC = @{}
    $QueryDC -split '\r?\n' | ForEach-Object {
        if ($_ -match '^(?<key>[^:]+):\s+(?<value>.+)$') {
            $OutDC[$Matches['key']] = $Matches['value']
        }
    }
    $LocalDC = New-Object PSObject -Property $OutDC
    $FinalDC = $LocalDC.'           DC' -replace "\\",""
    $FinalDC = $FinalDC -replace ".[\w]{2}.[\w]{2}.[\w]{5}.[\w]{3}"
    Logger -FilePath $LogFileLoc -Text "Resetting Computer Machine Password." -OPWindow
    Reset-ComputerMachinePassword -Server $FinalDC -Credential $creds
    Start-Sleep -Seconds 2
    Logger -FilePath $LogFileLoc -Text "Rebooting computer." -OPWindow
    Restart-Computer -Force
}
function CAC-Removal{
    Param(
        [switch]$enable,
        [switch]$disable
        )
    $outputBox.Text = ''
    if($disable.ispresent){
        Logger -FilePath $LogFileLoc -Text "Setting the Smart Card Removal Policy Service to Automatic startup" -OPWindow
        Set-Service -Name "SCPolicySvc" -StartupType Automatic
        Start-Service "SCPolicySvc"
        Logger -FilePath $LogFileLoc -Text "Smart Card Removal Policy Service has been started" -OPWindow
    }
    if($enable.ispresent){
        Logger -FilePath $LogFileLoc -Text "Stopping SCPolicySvc Service"
        Stop-Service "SCPolicySvc"
        Logger -FilePath $LogFileLoc -Text "Smart Card Removal Policy Service has been stopped" -OPWindow
        Set-Service -Name "SCPolicySvc" -StartupType Disabled
        Logger -FilePath $LogFileLoc -Text "Smart Card Removal Policy Service has been disabled to prevent auto starting." -OPWindow
    }
    Logger -FilePath $LogFileLoc -Text "Completed" -OPWindow
}
function GP-Update{
    Logger -FilePath $LogFileLoc -Text "Running GPUpdate" -OPWindow
    gpupdate
    $GPC = Get-EventLog -ComputerName $env:COMPUTERNAME -LogName System -InstanceId 1502 -Newest 1 
    $outputBox.AppendText($GPC.TimeGenerated)
    $outputBox.AppendText($GPC.Message)
    $outputBox.AppendText("`n")
    $outputBox.AppendText("`n")
    $GPU = Get-EventLog -ComputerName $env:COMPUTERNAME -LogName System -InstanceId 1501 -Newest 1 
    $outputBox.AppendText($GPU.TimeGenerated)
    $outputBox.AppendText($GPU.Message)
    $outputBox.AppendText("`n")
    $outputBox.AppendText("`n")
    $GPE = Get-EventLog -ComputerName $env:COMPUTERNAME -LogName System -InstanceId 1006 -Newest 1
    $outputBox.AppendText($GPE.TimeGenerated)
    $outputBox.AppendText($GPE.Message)
    $outputBox.AppendText("`n")
    $outputBox.AppendText("`n")
    $outputBox.AppendText("Completed")
    Logger -FilePath $LogFileLoc -Text "Completed.  Results in System Event Logs"
}
function ServerInfo{
    $outputBox.Text =''
    Logger -FilePath $LogFileLoc -Text "Getting site info" -OPWindow
    Try{
        Import-Module ActiveDirectory
        $homeDC = Get-ADDomainController -Discover
        $DNSServers = Get-DnsClientServerAddress -InterfaceAlias Ethernet -AddressFamily IPv4
        $HDC = $homeDC.Name
        $HDCS = $homeDC.Site
        Logger -FilePath $LogFileLoc -Text "Closest discovered Domain Controller is: $HDC on the $HDCS site" -OPWindow
        Logger -FilePath $LogFileLoc -Text "If this is not accurate, or consistently wrong, sites and services may be misconfigured, or there are server problems." -OPWindow -NewLine
        Logger -FilePath $LogFileLoc -Text "Your DNS Servers are: $($DNSServers.ServerAddresses)" -OPWindow -NewLine
        $mailtest = Test-Connection -ComputerName mail.ds.dodea.edu -Count 1
        $mailIPs = Resolve-DnsName -Name mail.ds.dodea.edu 
        Logger -FilePath $LogFileLoc -Text "E-Mail Web addresses are: $($mailIPs.IP4Address)" -OPWindow
        if($mailtest.ResponseTime -ne $null)
        {
            Logger -FilePath $LogFileLoc -Text "Email connectivity to:$($mailtest.Address) - $($mailtest.IPV4Address) - is reachable." -OPWindow -NewLine
        }
        else
        {
            Logger -FilePath $LogFileLoc -Text "Email connectivity to:  $($mailtest.Address) - $($mailtest.IPV4Address) - could not be reached." -OPWindow -NewLine
        }
        $matchme = $homeDC.Name -match "\w+-\w+"
        $hostname = $Matches[0]
        $hostname = $hostname + "*"
        $serverlist = Get-ADComputer -Filter {Name -like $hostname} -SearchScope Subtree | select Name
        $outputBox.AppendText("`n")
        Logger -FilePath $LogFileLoc -Text "Creating jobs to ping multiple computers"
        Foreach($server in $serverlist){
            $matchme = $server -match "\w+-\w+-\w+\d+"
            $server = $Matches.Values
            Set-Variable -Name "Status_$server" -Value (Test-Connection -ComputerName $server -AsJob -Count 1)
        }
        #check the results of each ping job
        Logger -FilePath $LogFileLoc -Text "Fetching results of ping jobs"
        Get-Variable "Status_*" -ValueOnly | Foreach {
            $Status = Wait-Job $_ | Receive-Job
            if ($Status.ResponseTime -ne $null ) {
                Logger -FilePath $LogFileLoc -Text "$($Status.Address)-$($Status.IPV4Address) is reachable." -OPWindow -NewLine
            }
            else{
                Logger -FilePath $LogFileLoc -Text "$($Status.Address)-$($Status.IPV4Address) could not be reached." -OPWindow -NewLine
            }
        }
        Logger -FilePath $LogFileLoc -Text "Completed querying servers" -OPWindow
        Remove-Job *
    }
    Catch{
        Logger -FilePath $LogFileLoc -Text "This function requires the ActiveDirectory Module for PowerShell to be installed.  Please ensure that the RSAT Tools have been install and then under Programs and Features, click Turn Windows features on or off.  Under Remote Server Administration Tools, open Role Administration Tools, open AD DS and AD LDS Tools.  Ensure that Active Directory Module for Windows PowerShell is selected." -OPWindow
    }
}
function ComputerCleanup{
    $outputBox.Text = ''
    Logger -FilePath $LogFileLoc -Text "Attempting to load Active Directory PowerShell module" -OPWindow
    Try{
        Import-Module ActiveDirectory
        Logger -FilePath $LogFileLoc -Text "Searching the Computers OU for matching computer names..." -OPWindow
        $complist = Get-ADComputer -Filter ("(Name -like 'STHS*' -or Name -like 'STES*' -or Name -like 'PAES*' -or Name -like 'PAMS*' -or Name -like 'ROBI*') -and lastLogonTimestamp -gt 0") -SearchBase 'CN=Computers,DC=eu,DC=DS,DC=DODEA,DC=EDU' -SearchScope OneLevel -Properties Name,DistinguishedName,OperatingSystem
        if($complist -ne $null){
            Logger -FilePath $LogFileLoc -Text "Found matching computer objects, moving them to the proper OU now..." -OPWindow
            Foreach($comp in $complist){  
                if($comp.OperatingSystem -like 'Windows 7 Enterprise'){
                    if($comp.Name -like 'STHS*'){$comp | Move-ADObject -TargetPath 'OU=Windows,OU=Computers,OU=Stuttgart HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                    if($comp.Name -like 'STES*'){$comp | Move-ADObject -TargetPath 'OU=Windows,OU=Computers,OU=Stuttgart ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                    if($comp.Name -like 'PAES*'){$comp | Move-ADObject -TargetPath 'OU=Windows,OU=Computers,OU=Patch ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                    if($comp.Name -like 'PAMS*'){$comp | Move-ADObject -TargetPath 'OU=Windows,OU=Computers,OU=Patch MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                    if($comp.Name -like 'ROBI*'){$comp | Move-ADObject -TargetPath 'OU=Windows,OU=Computers,OU=Robinson Barracks ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                }
                if($comp.OperatingSystem -like 'Windows 10 Enterprise'){
                    if($comp.Name -like 'STHS*'){$comp | Move-ADObject -TargetPath 'OU=Windows10,OU=Computers,OU=Stuttgart HS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                    if($comp.Name -like 'STES*'){$comp | Move-ADObject -TargetPath 'OU=Windows10,OU=Computers,OU=Stuttgart ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                    if($comp.Name -like 'PAES*'){$comp | Move-ADObject -TargetPath 'OU=Windows10,OU=Computers,OU=Patch ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                    if($comp.Name -like 'PAMS*'){$comp | Move-ADObject -TargetPath 'OU=Windows10,OU=Computers,OU=Patch MS,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                    if($comp.Name -like 'ROBI*'){$comp | Move-ADObject -TargetPath 'OU=Windows10,OU=Computers,OU=Robinson Barracks ES,OU=Schools and Offices,DC=eu,DC=DS,DC=DODEA,DC=EDU'}
                }
            }
        }
        else{Logger -FilePath $LogFileLoc -Text "Found no matching computer objects." -OPWindow}
    }
    Catch{
        Logger -FilePath $LogFileLoc -Text "This function requires the ActiveDirectory Module for PowerShell to be installed.  Please ensure that the RSAT Tools have been install and then under Programs and Features, click Turn Windows features on or off.  Under Remote Server Administration Tools, open Role Administration Tools, open AD DS and AD LDS Tools.  Ensure that Active Directory Module for Windows PowerShell is selected." -OPWindow
    }
}
function DriveMappings{
    $outputBox.Text = ''
    Logger -FilePath $LogFileLoc -Text "Matching School location" -OPWindow
    $SchoolMatcher = SchoolMatcher -CompName "$env:COMPUTERNAME"
    Logger -FilePath $LogFileLoc -Text "Grabbing currently logged on user information..." -OPWindow
    Logger -FilePath $LogFileLoc -Text "Querying AD for user H: drive" -OPWindow
    #$ADUser = directorysearcher -Filter "SamAccountName=$CurrentUserName" -SearchScope subtree -SearchType FindOne -Domain $CurrentUserDomain
    $DirSerLoc = ($SchoolMatcher.OU -replace "OU=Production,OU=Computers,","")
    $ADUser = directorysearcher -SearchScope subtree -SearchType FindAll -Filter "SamAccountName=$CurrentUserName" -SearchLocation $DirSerLoc
    $ADUserHome = $ADUser.properties.homedirectory[0]
    Logger -FilePath $LogFileLoc -Text "Building batch file for the user" -OPWindow
    Logger -FilePath $LogFileLoc -Text "Building net use commands"
    $DriveH = "net use H: $ADUserHome /persistent:Yes"
    $command = "nltest"
    $args = "/DSGETDC:"
    $QueryDC = & $command $args 
    $OutDC = @{}
    $QueryDC -split '\r?\n' | ForEach-Object {
        if ($_ -match '^(?<key>[^:]+):\s+(?<value>.+)$') {
            $OutDC[$Matches['key']] = $Matches['value']
        }
    }
    $LocalDC   = New-Object PSObject -Property $OutDC
    if($SchoolMatcher.MDrive)   {$DriveM = "net use M: $($SchoolMatcher.MDrive) /persistent:Yes"    }
    if($SchoolMatcher.IDrive)   {$DriveI = "net use I: $($SchoolMatcher.IDrive) /persistent:Yes"   }
    if($SchoolMatcher.HomeDrive){$DriveO = "net use O: $($SchoolMatcher.HomeDrive) /persistent:Yes"}
    Logger -FilePath $LogFileLoc -Text "Creating MapDrive.bat on loggedon users desktop" -OPWindow
    New-Item -Path "C:\Users\$CurrentUserName\Desktop\" -Name MapDrives.bat -ItemType File -Value $DriveH -Force
    "`n" | Out-File -FilePath "C:\Users\$CurrentUserName\Desktop\MapDrives.bat" -Append -Encoding ascii
    if($DriveM -ne $null){$DriveM | Out-File -FilePath "C:\Users\$CurrentUserName\Desktop\MapDrives.bat" -Append -Encoding ascii} ; Logger -FilePath $LogFileLoc -Text "Discovered MDrive, adding to batch file"
    if($DriveI -ne $null){$DriveI | Out-File -FilePath "C:\Users\$CurrentUserName\Desktop\MapDrives.bat" -Append -Encoding ascii} ; Logger -FilePath $LogFileLoc -Text "Discovered IDrive, adding to batch file"
    if($DriveO -ne $null){$DriveO | Out-File -FilePath "C:\Users\$CurrentUserName\Desktop\MapDrives.bat" -Append -Encoding ascii} ; Logger -FilePath $LogFileLoc -Text "Discovered ODrive, adding to batch file"
    Logger -FilePath $LogFileLoc -Text "Completed creating the file.  User should be able to run the batch file." -OPWindow
}
function BackupUserFiles{}#robocopy shit WIP
function GenerateCSV{
    Logger -FilePath $LogFileLoc -Text "Grabbing computername to see which site the user is at" -OPWindow
    $SchoolMatcher = SchoolMatcher -CompName "$env:COMPUTERNAME"
    Logger -FilePath $LogFileLoc -Text "Prompting user for domain selection" -OPWindow
    $Form.TopMost = $false
    do {
        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
        $DomainInput = [Microsoft.VisualBasic.Interaction]::InputBox("Please input EU or DS to pick which domain you want to search.", "Domain Selection","")
        $matchme = $DomainInput -match "EU|DS"
    
    } until ($matchme -eq $true)
    $Form.TopMost = $true
    if($Matches.Values -eq 'EU'){$DirSerLoc = $SchoolMatcher.OldOU}
    if($Matches.Values -eq 'DS'){$DirSerLoc = $SchoolMatcher.OU}
    Logger -FilePath $LogFileLoc -Text "Searching Active Directory for Windows 10 computers" -OPWindow
    $list = directorysearcher -SearchScope subtree -SearchType FindAll -Filter "(&(objectcategory=computer)(operatingsystem=*Windows 10 Enterprise*))" -SearchLocation $DirSerLoc
    $op = '' | Select Name
    Logger -FilePath $LogFileLoc -Text "Generating output of Win10 computers" -OPWindow
    foreach($computer in $list)
    {
        $op.Name = $computer.properties.name[0]
        $op
    }
    Logger -FilePath $LogFileLoc -Text "Exiting GenerateCSV function" -OPWindow
}
#endregion Work Functions

function About {
    # About Form Objects
    $aboutForm             = New-Object System.Windows.Forms.Form
    $aboutFormExit         = New-Object System.Windows.Forms.Button
    $aboutFormImage        = New-Object System.Windows.Forms.PictureBox
    $SCCMNameLabel         = New-Object System.Windows.Forms.Label
    $SCCMClientLabel       = New-Object System.Windows.Forms.Label
    $SCCMSiteLabel         = New-Object System.Windows.Forms.Label
    $SCCMSiteLabel1        = New-Object System.Windows.Forms.Label
    $aboutFormText         = New-Object System.Windows.Forms.Label
    $RunClientActionLabel  = New-Object System.Windows.Forms.Label
    $RunClientActionLabel1 = New-Object System.Windows.Forms.Label
    $BackupLabel           = New-Object System.Windows.Forms.Label
    $BackupLabel1          = New-Object System.Windows.Forms.Label
    $gpupdatelabel         = New-Object System.Windows.Forms.Label
    $gpupdatelabel1        = New-Object System.Windows.Forms.Label
    $serverinfolabel       = New-Object System.Windows.Forms.Label
    $serverinfolabel1      = New-Object System.Windows.Forms.Label
    $userdrivemaplabel     = New-Object System.Windows.Forms.Label
    $userdrivemaplabel1    = New-Object System.Windows.Forms.Label
    $trustRlabel           = New-Object System.Windows.Forms.Label
    $trustRlabel1          = New-Object System.Windows.Forms.Label
    $CACpolicylabel        = New-Object System.Windows.Forms.Label
    $CACpolicylabel1       = New-Object System.Windows.Forms.Label
    $informationlabel      = New-Object System.Windows.Forms.Label
    $CleanCompLabel        = New-Object System.Windows.Forms.Label
    $CleanCompLabel1       = New-Object System.Windows.Forms.Label
    $renamecomputerslabel  = New-Object System.Windows.Forms.Label
    $renamecomputerslabel1 = New-Object System.Windows.Forms.Label

    # About Form
    $aboutForm.AcceptButton  = $aboutFormExit
    $aboutForm.CancelButton  = $aboutFormExit
    $aboutForm.ClientSize    = "800, 800"
    $aboutForm.ControlBox    = $false
    $aboutForm.ShowInTaskBar = $false
    $aboutForm.TopMost       = $true
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.Text          = "Command info"
    $aboutForm.Add_Load($aboutForm_Load)

    # About SCCM Name Label
    $SCCMNameLabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $SCCMNameLabel.Location = "50, 20"
    $SCCMNameLabel.Size     = "200, 18"
    $SCCMNameLabel.Text     = "Reinstall Client"
    $aboutForm.Controls.Add($SCCMNameLabel)

    # About SCCM Client Label
    $SCCMClientLabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $SCCMClientLabel.Location = "10, 40"
    $SCCMClientLabel.AutoSize = $true
    $SCCMClientLabel.Text     = "This will uninstall the current SCCM Client, 
and re-install the client and then automatically
try and assign the site code and run client actions."
    $aboutForm.Controls.Add($SCCMClientLabel)

    # About Check Site Label
    $SCCMSiteLabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $SCCMSiteLabel.Location = "50, 100"
    $SCCMSiteLabel.Size     = "200, 18"
    $SCCMSiteLabel.Text     = "Check Site"
    $aboutForm.Controls.Add($SCCMSiteLabel)

    # About Check Site Label
    $SCCMSiteLabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $SCCMSiteLabel1.Location = "10, 120"
    $SCCMSiteLabel1.AutoSize = $true
    $SCCMSiteLabel1.Text     = "This will check the SCCM Site on the client,
and attempt to set it to E01 automatically."
    $aboutForm.Controls.Add($SCCMSiteLabel1)

    # About Run Client Action Label
    $RunClientActionLabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $RunClientActionLabel.Location = "50, 160"
    $RunClientActionLabel.Size     = "200, 18"
    $RunClientActionLabel.Text     = "Run Client Actions"
    $aboutForm.Controls.Add($RunClientActionLabel)

    # About Run Client Action Label
    $RunClientActionLabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $RunClientActionLabel1.Location = "10, 180"
    $RunClientActionLabel1.AutoSize = $true
    $RunClientActionLabel1.Text     = "This will automatically run the actions
on the SCCM Client Actions tab."
    $aboutForm.Controls.Add($RunClientActionLabel1)

    # About Backup Label
    $BackupLabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $BackupLabel.Location = "50, 220"
    $BackupLabel.Size     = "200, 18"
    $BackupLabel.Text     = "Backup user files"
    $aboutForm.Controls.Add($BackupLabel)

    # About Run Client Action Label
    $BackupLabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $BackupLabel1.Location = "10, 240"
    $BackupLabel1.AutoSize = $true
    $BackupLabel1.Text     = "This backs up all the files for the currently
logged in user."
    $aboutForm.Controls.Add($BackupLabel1)

    # About Gpupdate Label
    $gpupdatelabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $gpupdatelabel.Location = "50, 280"
    $gpupdatelabel.Size     = "200, 18"
    $gpupdatelabel.Text     = "Group Policy Update"
    $aboutForm.Controls.Add($gpupdatelabel)

    # About Gpupdate Label
    $gpupdatelabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $gpupdatelabel1.Location = "10, 300"
    $gpupdatelabel1.AutoSize = $true
    $gpupdatelabel1.Text     = "Runs a gpupdate and displays results."
    $aboutForm.Controls.Add($gpupdatelabel1)

    # About Server Info Label
    $serverinfolabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $serverinfolabel.Location = "50, 320"
    $serverinfolabel.Size     = "200, 18"
    $serverinfolabel.Text     = "Get Server Info"
    $aboutForm.Controls.Add($serverinfolabel)

    # About Server Info Label
    $serverinfolabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $serverinfolabel1.Location = "10, 340"
    $serverinfolabel1.AutoSize = $true
    $serverinfolabel1.Text     = "Discovers closest Domain Controller
and uses that to test connectivity to 
nearby servers."
    $aboutForm.Controls.Add($serverinfolabel1)

    # About User Drive Mapping Label
    $userdrivemaplabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $userdrivemaplabel.Location = "50, 400"
    $userdrivemaplabel.Size     = "200, 18"
    $userdrivemaplabel.Text     = "User Drive Mapping"
    $aboutForm.Controls.Add($userdrivemaplabel)

    # About User Drive Mapping Label
    $userdrivemaplabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $userdrivemaplabel1.Location = "10, 420"
    $userdrivemaplabel1.AutoSize = $true
    $userdrivemaplabel1.Text     = "This will check who the current
logged on user is and then queries AD to
find out their drive mapping."
    $aboutForm.Controls.Add($userdrivemaplabel1)

    # About Rename Computers Label
    $renamecomputerslabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $renamecomputerslabel.Location = "50, 400"
    $renamecomputerslabel.Size     = "200, 18"
    $renamecomputerslabel.Text     = "Rename Computers"
    $aboutForm.Controls.Add($renamecomputerslabel)

    # About Rename Computers Label
    $renamecomputerslabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $renamecomputerslabel1.Location = "10, 420"
    $renamecomputerslabel1.AutoSize = $true
    $renamecomputerslabel1.Text     = "This action will import a list of computers
from a .csv file and construct the new computer 
name and then attempt to remotely change the 
computers name.  This will reboot the renamed
computer."
    $aboutForm.Controls.Add($renamecomputerslabel1)

    # About Trust Relationship Label
    $trustRlabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $trustRlabel.Location = "50, 520"
    $trustRlabel.Size     = "200, 18"
    $trustRlabel.Text     = "Fix Trust Relationship"
    $aboutForm.Controls.Add($trustRlabel)

    # About Trust Relationship Label
    $trustRlabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $trustRlabel1.Location = "10, 540"
    $trustRlabel1.AutoSize = $true
    $trustRlabel1.Text     = "This will fix the local computer by reaching 
out and resetting the computer password 
with active directory and then rebooting
the machine."
    $aboutForm.Controls.Add($trustRlabel1)

    # About CAC Policy Label
    $CACpolicylabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $CACpolicylabel.Location = "400, 20"
    $CACpolicylabel.Size     = "200, 18"
    $CACpolicylabel.Text     = "CAC Removal Policy"
    $aboutForm.Controls.Add($CACpolicylabel)

    # About CAC Policy  Label
    $CACpolicylabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $CACpolicylabel1.Location = "360, 40"
    $CACpolicylabel1.AutoSize = $true
    $CACpolicylabel1.Text     = "This will start or stop the Smart Card Removal 
Service which will allow users to remove their
CAC and not lock the computer."
    $aboutForm.Controls.Add($CACpolicylabel1)

    # About Cleanup Computers Label
    $CleanCompLabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $CleanCompLabel.Location = "400, 100"
    $CleanCompLabel.Size     = "200, 18"
    $CleanCompLabel.Text     = "Clean Computer OU"
    $aboutForm.Controls.Add($CleanCompLabel)

    # About Cleanup Computers Label
    $CleanCompLabel1.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $CleanCompLabel1.Location = "360, 120"
    $CleanCompLabel1.AutoSize = $true
    $CleanCompLabel1.Text     = "This will query all the computers in the Computers OU 
and then move the computer to the correct OU
based on the computer name and operating system."
    $aboutForm.Controls.Add($CleanCompLabel1)

    # About Information Label
    $informationlabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $informationlabel.Location = "275, 700"
    $informationlabel.AutoSize = $true
    $informationlabel.Text     = "           Created and maintained by 
Jim Lower - james.lower.ctr@dodea.edu"
    $aboutForm.Controls.Add($informationlabel)

    # About Exit Button
    $aboutFormExit.Location = "375, 740"
    $aboutFormExit.Text     = "OK"
    $aboutForm.Controls.Add($aboutFormExit)

    [void]$aboutForm.ShowDialog()
}
function HelpfulLinks{
    Add-Type -AssemblyName System.Windows.Forms
    
    $helpfulLinksForm = New-Object system.Windows.Forms.Form
    $helpfulLinksForm.Size = New-Object System.Drawing.Size(500,200)
    $helpfulLinksForm.Icon = [System.IconExtractor]::Extract("shell32.dll", 13, $true)
    $helpfulLinksForm.Text = 'Helpful links'
    $helpfulLinksForm.StartPosition   = "CenterScreen"
    $helpfulLinksForm.BackColor = "#d1d1d1"
    $helpfulLinksForm.TopMost = $true
    
    $helpdesklinklabel = New-Object System.Windows.Forms.Label
    $helpdesklinklabel.Text = "Help Desk"
    $helpdesklinklabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $helpdesklinklabel.Location = New-Object System.Drawing.Size(5,20)
    $helpdesklinklabel.Size = New-Object System.Drawing.Size(100,20)
    $helpfulLinksForm.Controls.Add($helpdesklinklabel)
    
    $helpdesklinkTextBox = New-Object System.Windows.Forms.TextBox
    $helpdesklinkTextBox.ReadOnly = $true
    $helpdesklinkTextBox.Text = "https://help.ds.dodea.edu/support/staff/index.php?/Tickets/Manage/Index/user"
    $helpdesklinkTextBox.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $helpdesklinkTextBox.Location = New-Object System.Drawing.Size(5,40)
    $helpdesklinkTextBox.Size = New-Object System.Drawing.Size(450,50)
    $helpfulLinksForm.Controls.Add($helpdesklinkTextBox)
    
    $printserverlinklabel = New-Object System.Windows.Forms.Label
    $printserverlinklabel.Text = "Print Server"
    $printserverlinklabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $printserverlinklabel.Location = New-Object System.Drawing.Size(5,60)
    $printserverlinklabel.Size = New-Object System.Drawing.Size(100,20)
    $helpfulLinksForm.Controls.Add($printserverlinklabel)
    
    $printserverlinkTextBox = New-Object System.Windows.Forms.TextBox
    $printserverlinkTextBox.ReadOnly = $true
    $printserverlinkTextBox.Text = "http://dodea-directip.am.ds.dodea.edu/admin/?"
    $printserverlinkTextBox.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $printserverlinkTextBox.Location = New-Object System.Drawing.Size(5,80)
    $printserverlinkTextBox.Size = New-Object System.Drawing.Size(450,50)
    $helpfulLinksForm.Controls.Add($printserverlinkTextBox)
    
    $SolarWindslinklabel = New-Object System.Windows.Forms.Label
    $SolarWindslinklabel.Text = "Solar Winds"
    $SolarWindslinklabel.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $SolarWindslinklabel.Location = New-Object System.Drawing.Size(5,100)
    $SolarWindslinklabel.Size = New-Object System.Drawing.Size(100,20)
    $helpfulLinksForm.Controls.Add($SolarWindslinklabel)
    
    $SolarWindslinkTextBox = New-Object System.Windows.Forms.TextBox
    $SolarWindslinkTextBox.ReadOnly = $true
    $SolarWindslinkTextBox.Text = "https://solarwinds.ds.dodea.edu"
    $SolarWindslinkTextBox.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $SolarWindslinkTextBox.Location = New-Object System.Drawing.Size(5,120)
    $SolarWindslinkTextBox.Size = New-Object System.Drawing.Size(450,50)
    $helpfulLinksForm.Controls.Add($SolarWindslinkTextBox)
    
    $helpfulLinksForm.ShowDialog()
}

function Generate-Form {
    #region Building the form
    #Logger -FilePath $LogFileLoc -Text "Script has been elevated, opening Form"
    # Create Icon Extractor Assembly
    $code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@
    Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing
    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing
    [Windows.Forms.Application]::EnableVisualStyles()

    #Settings up form objects
    #####################################################################
    $Form             = New-Object System.Windows.Forms.Form
    $menuMain         = New-Object System.Windows.Forms.MenuStrip
    $menuFile         = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuMigration    = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuRenameComp   = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuDomainChange = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuExit         = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuLinks        = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuHelp         = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuAbout        = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuFixWireless  = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuGenCSV       = New-Object System.Windows.Forms.ToolStripMenuItem

    $Form.Text = "DoDEA Service Desk Tool v2.5"
    $Form.BackColor = "#d1d1d1"
    $Form.StartPosition   = "CenterScreen"
    $Form.add_FormClosing({
        Logger -FilePath $LogFileLoc -Text "Registered Closing event"
        try{
            Logger -FilePath $LogFileLoc -Text "Attempting to cleanup old files"
            Remove-Item -Path "C:\Users\$CurrentUserName\ServiceDeskLogFile.txt" -Force | Out-Null
            Remove-Item "C:\Users\$CurrentUserName\RenameRemoteComputers.bat" -Force
            Logger -FilePath $LogFileLoc -Text "Cleanup complete"
        }catch{}
        Logger -FilePath $LogFileLoc -Text "Closing script"
    })
    $Form.Size = New-Object System.Drawing.Size(600,600)
    $Form.TopMost = $true
    $Form.Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)
    #endregion Building the form

    #region Menus
    # Menu Options - File / Exit
    #####################################################################
    [void]$Form.Controls.Add($menuMain)
    $menuFile.Text = "&File"
    $menuFile.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    [void]$menuMain.Items.Add($menuFile)

    $menuExit.Image        = [System.IconExtractor]::Extract("shell32.dll", 27, $true)
    $menuExit.ShortcutKeys = "Control, X"
    $menuExit.Text         = "&Exit"
    $menuExit.Add_Click({
        Logger -FilePath $LogFileLoc -Text "Registered Closing event"
        try{
            Logger -FilePath $LogFileLoc -Text "Attempting to clean up old files"
            Remove-Item -Path "C:\Users\$CurrentUserName\ServiceDeskLogFile.txt" -Force | Out-Null
            Remove-Item "C:\Users\$CurrentUserName\RenameRemoteComputers.bat" -Force
            Logger -FilePath $LogFileLoc -Text "Completed cleaning"
            }catch{}
        Logger -FilePath $LogFileLoc -Text "Closing script"
        $Form.Close()
    })
    [void]$menuFile.DropDownItems.Add($menuExit)

    # Menu Options - Migration Tools
    #####################################################################
    $menuMigration.Text = "&Migration"
    $menuMigration.Font     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    [void]$menuMain.Items.Add($menuMigration)

    $menuDomainChange.Image        = [System.IconExtractor]::Extract("shell32.dll", 42, $true)
    $menuDomainChange.Text         = "&Change Domains"
    $menuDomainChange.Add_Click({MigrateComputers -ChangeDomain  ; Logger -FilePath $LogFileLoc -Text "Calling MigrateComuters ChangeDomain function"})
    [void]$menuMigration.DropDownItems.Add($menuDomainChange)

    $menuRenameComp.Image        = [System.IconExtractor]::Extract("shell32.dll", 94, $true)
    $menuRenameComp.Text         = "&Rename Computers"
    $menuRenameComp.Add_Click({MigrateComputers -RenameComputers ; Logger -FilePath $LogFileLoc -Text "Calling MigrateComputers RenameComputers function"})
    [void]$menuMigration.DropDownItems.Add($menuRenameComp)

    $menuFixWireless.Image        = [System.IconExtractor]::Extract("shell32.dll", 248, $true)
    $menuFixWireless.Text         = "&Fix Wireless"
    $menuFixWireless.Add_Click({MigrateComputers -FixWireless    ; Logger -FilePath $LogFileLoc -Text "Calling MigrateComputers FixWireless function"})
    [void]$menuMigration.DropDownItems.Add($menuFixWireless)

    $menuGenCSV.Image        = [System.IconExtractor]::Extract("shell32.dll", 282, $true)
    $menuGenCSV.Text         = "&Generate CSV of computers"
    $menuGenCSV.Add_Click({Logger -FilePath $LogFileLoc -Text "Calling GenerateCSV function" ; GenerateCSV | Export-Csv -NoTypeInformation -Path "C:\Users\$CurrentUserName\Desktop\ComputerQuery.csv"    ; Logger -FilePath $LogFileLoc -Text "Exporting ComputerQuery CSV to the desktop" -OPWindow})
    [void]$menuMigration.DropDownItems.Add($menuGenCSV)

    # Menu Options - Help
    #####################################################################
    $menuHelp.Text                     = "&Help"
    $menuHelp.Font                     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    [void]$menuMain.Items.Add($menuHelp)

    $menuLinks.Image                   = [System.IconExtractor]::Extract("shell32.dll", 13, $true)
    $menuLinks.Text                    = "&Links"
    $menuLinks.Add_Click({HelpfulLinks})
    [void]$menuHelp.DropDownItems.Add($menuLinks)
    
    $menuAbout.Image                   = [System.Drawing.SystemIcons]::Information
    $menuAbout.Text                    = "About"
    $menuAbout.Add_Click({$Form.TopMost = $false;About})
    [void]$menuHelp.DropDownItems.Add($menuAbout)
    #endregion Menus
    #region Text and Buttons
    #Adds text instructions for the user
    #####################################################################
    $SCCMLabel                         = New-Object System.Windows.Forms.Label
    $SCCMLabel.Text                    = "SCCM Tools"
    $SCCMLabel.Font                    = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $SCCMLabel.Location                = New-Object System.Drawing.Size(70,27)
    $SCCMLabel.AutoSize                = $True
    $Form.Controls.Add($SCCMLabel)

    $ToolsLabel                        = New-Object System.Windows.Forms.Label
    $ToolsLabel.Text                   = "Misc Tools"
    $ToolsLabel.Font                   = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $ToolsLabel.Location               = New-Object System.Drawing.Size(250,27)
    $ToolsLabel.AutoSize               = $True
    $Form.Controls.Add($ToolsLabel)

    $FixLabel                          = New-Object System.Windows.Forms.Label
    $FixLabel.Text                     = "Fix actions"
    $FixLabel.Font                     = New-Object Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $FixLabel.Location                 = New-Object System.Drawing.Size(430,27)
    $FixLabel.AutoSize                 = $True
    $Form.Controls.Add($FixLabel)

    # Add Buttons
    #####################################################################
    $reinstallB                        = New-Object System.Windows.Forms.Button
    $reinstallB.Location               = New-Object System.Drawing.Size(25,50)
    $reinstallB.Size                   = New-Object System.Drawing.Size(175,23)
    $reinstallB.Text                   = "Reinstall client"
    $reinstallB.BackColor              = "#FFFFFF"
    $Form.Controls.Add($reinstallB)

    $siteB                             = New-Object System.Windows.Forms.Button
    $siteB.Location                    = New-Object System.Drawing.Size(25,100)
    $siteB.Size                        = New-Object System.Drawing.Size(175,23)
    $siteB.Text                        = "Check Site "
    $siteB.BackColor                   = "#FFFFFF"
    $Form.Controls.Add($siteB)

    $runactionsB                       = New-Object System.Windows.Forms.Button
    $runactionsB.Location              = New-Object System.Drawing.Size(25,150)
    $runactionsB.Size                  = New-Object System.Drawing.Size(175,23)
    $runactionsB.Text                  = "Run Client actions "
    $runactionsB.BackColor             = "#FFFFFF"
    $Form.Controls.Add($runactionsB)

    $BackupUser                        = New-Object System.Windows.Forms.Button
    $BackupUser.Location               = New-Object System.Drawing.Size(200,50)
    $BackupUser.Size                   = New-Object System.Drawing.Size(175,23)
    $BackupUser.Text                   = "Backup user files "
    $BackupUser.BackColor              = "#FFFFFF"
    $Form.Controls.Add($BackupUser)

    $gpupdate                          = New-Object System.Windows.Forms.Button
    $gpupdate.Location                 = New-Object System.Drawing.Size(200,100)
    $gpupdate.Size                     = New-Object System.Drawing.Size(175,23)
    $gpupdate.Text                     = "Group Policy Update "
    $gpupdate.BackColor                = "#FFFFFF"
    $Form.Controls.Add($gpupdate)

    $serverinfo                        = New-Object System.Windows.Forms.Button
    $serverinfo.Location               = New-Object System.Drawing.Size(200,150)
    $serverinfo.Size                   = New-Object System.Drawing.Size(175,23)
    $serverinfo.Text                   = "Get server info"
    $serverinfo.BackColor              = "#FFFFFF"
    $Form.Controls.Add($serverinfo)

    $drivemapper                       = New-Object System.Windows.Forms.Button
    $drivemapper.Location              = New-Object System.Drawing.Size(200,200)
    $drivemapper.Size                  = New-Object System.Drawing.Size(175,23)
    $drivemapper.Text                  = "User Drive Mapping"
    $drivemapper.BackColor             = "#FFFFFF"
    $Form.Controls.Add($drivemapper)

    $resetcomppw                       = New-Object System.Windows.Forms.Button
    $resetcomppw.Location              = New-Object System.Drawing.Size(390,50)
    $resetcomppw.Size                  = New-Object System.Drawing.Size(175,23)
    $resetcomppw.Text                  = "Fix Trust Relationship"
    $resetcomppw.BackColor             = "#FFFFFF"
    $Form.Controls.Add($resetcomppw)

    $cacpolicy                         = New-Object System.Windows.Forms.Button
    $cacpolicy.Location                = New-Object System.Drawing.Size(390,100)
    $cacpolicy.Size                    = New-Object System.Drawing.Size(175,23)
    $cacpolicy.Text                    = "CAC Removal Policy"
    $cacpolicy.BackColor               = "#FFFFFF"
    $Form.Controls.Add($cacpolicy)

    $EnableRButton                     = New-Object system.Windows.Forms.RadioButton
    $EnableRButton.text                = "Enable"
    $EnableRButton.AutoSize            = $true
    $EnableRButton.width               = 104
    $EnableRButton.height              = 20
    $EnableRButton.location            = New-Object System.Drawing.Point(410,120)
    $EnableRButton.Font                = 'Microsoft Sans Serif,9'
    
    $DisableRButton                    = New-Object system.Windows.Forms.RadioButton
    $DisableRButton.text               = "Disable"
    $DisableRButton.AutoSize           = $true
    $DisableRButton.width              = 104
    $DisableRButton.height             = 20
    $DisableRButton.location           = New-Object System.Drawing.Point(490,120)
    $DisableRButton.Font               = 'Microsoft Sans Serif,9'

    $Form.controls.AddRange(@($EnableRButton,$DisableRButton))

    $computercleanup                   = New-Object System.Windows.Forms.Button
    $computercleanup.Location          = New-Object System.Drawing.Size(390,150)
    $computercleanup.Size              = New-Object System.Drawing.Size(175,23)
    $computercleanup.Text              = "Clean Computer OU"
    $computercleanup.BackColor         = "#FFFFFF"
    $Form.Controls.Add($computercleanup)

    $outputBox                         = New-Object System.Windows.Forms.RichTextBox
    $outputBox.Location                = New-Object System.Drawing.Size(10,350)
    $outputBox.Size                    = New-Object System.Drawing.Size(565,200)
    $outputBox.Font                    = New-Object Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular)
    $outputBox.MultiLine               = $True
    $outputBox.ScrollBars              = "Vertical"
    $outputBox.ReadOnly                = $True
    $Form.Controls.Add($outputBox)
    #endregion Text and Buttons
    #region Actions
    #Tells the buttons what to do when clicked
    #####################################################################
    $reinstallB.Add_Click({SCCM-Repair -Reinstall   ; Logger -FilePath $LogFileLoc -Text "Calling SCCM-Repair Reinstall function"})
    $siteB.Add_Click({SCCM-Repair -Site             ; Logger -FilePath $LogFileLoc -Text "Calling SCCM-Repair Site function"})
    $runactionsB.Add_Click({SCCM-Repair -RunActions ; Logger -FilePath $LogFileLoc -Text "Calling SCCM-Repair RunActions function"})
    $resetcomppw.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType    = [System.Windows.MessageBoxButton]::YesNo
        $MessageIcon   = [System.Windows.MessageBoxImage]::Error
        $MessageBody   = "This resets the local Computer PW and requires a reboot.  This will fix the 'Trust Relationship' error message on the local computer.  Do you want to continue?"
        $MessageTitle  = "Confirm"
        $Result        = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
        Logger -FilePath $LogFileLoc -Text "Prompting user about the Reset-MachinePW function" -OPWindow
        if($Result -eq 'Yes'){Reset-MachinePW ; Logger -FilePath $LogFileLoc -Text "Calling Reset-MachinePW function"}    
    })
    $cacpolicy.Add_Click({
        if($global:CACPolicyRButton -eq ""){
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            $ButtonType  = [System.Windows.MessageBoxButton]::OK
            $MessageIcon = [System.Windows.MessageBoxImage]::Error
            $Result      =[System.Windows.Forms.Messagebox]::Show("Please select Enable or Disable first","CAC Removal Policy",$ButtonType,$MessageIcon)
            Logger -FilePath $LogFileLoc -Text "Prompting user to select a CAC Enable Radio Button" -OPWindow
        }
        if($global:CACPolicyRButton -eq "Enable"){CAC-Removal -enable   ; Logger -FilePath $LogFileLoc -Text "Calling CAC-Removal Enable function"}
        if($global:CACPolicyRButton -eq "Disable"){CAC-Removal -disable ; Logger -FilePath $LogFileLoc -Text "Calling CAC-Removal Disable function"}
    })
    $BackupUser.Add_Click({
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType  = [System.Windows.MessageBoxButton]::OK
        $MessageIcon = [System.Windows.MessageBoxImage]::Information
        $Result      =[System.Windows.Forms.Messagebox]::Show("This functionality has not been completed yet.","Backup user files",$ButtonType,$MessageIcon)
        Logger -FilePath $LogFileLoc -Text "Calling BackupUserFiles function"
        #BackupUserFiles
    })
    $gpupdate.Add_Click({GP-Update})
    $serverinfo.Add_Click({ServerInfo})
    $computercleanup.Add_Click({
        Logger -FilePath $LogFileLoc -Text "Calling ComputerCleanup function"
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType  = [System.Windows.MessageBoxButton]::OK
        $MessageIcon = [System.Windows.MessageBoxImage]::Information
        $Result      = [System.Windows.Forms.Messagebox]::Show("This functionality has not been completed yet.","Computer Cleanup",$ButtonType,$MessageIcon)
        #Add-Type -AssemblyName PresentationCore,PresentationFramework
        #$ButtonType = [System.Windows.MessageBoxButton]::YesNo
        #$MessageIcon = [System.Windows.MessageBoxImage]::Error
        #$MessageBody = "This will move Computer Objects out of the Computers OU based off of the Computer Name and Operating System.  Do you want to continue?"
        #$MessageTitle = "Confirm"
        #$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
        #if($Result -eq 'Yes'){ComputerCleanup}    
    })
    $drivemapper.Add_Click({DriveMappings ; Logger -FilePath $LogFileLoc -Text "Calling DriveMappings function"})
    $EnableRButton.Add_Click({$global:CACPolicyRButton = "Enable"   ;Write-Host $CACPolicyRButton ; Logger -FilePath $LogFileLoc -Text "Enable Radio Button"})
    $DisableRButton.Add_Click({$global:CACPolicyRButton = "Disable" ;Write-Host $CACPolicyRButton ; Logger -FilePath $LogFileLoc -Text "Disable Radio Button"})
    #endregion Actions
    #Show the Form 
    #####################################################################
    Hide-Console
    $form.ShowDialog()
    
}

#region Run Form
if ((Test-Path "C:\Users\$CurrentUserName\ElevateScript.txt") -eq $true){Generate-Form | Out-Null}
else{
    New-Item -Path "C:\Users\$CurrentUserName\" -Name ElevateScript.txt -ItemType File -Value $currentdate -Force | Out-Null
    Start-Process -FilePath "powershell" -Verb runAs -WorkingDirectory $env:windir -ArgumentList "-file `"$($PSScriptRoot)`"\ServiceDesk.ps1"
}
#endregion Run Form