<#
.SYNOPSIS
    Install PowerCLI Offline
.DESCRIPTION
    Install the nuget package manamgment prereq then installs PowerCLI module
.PARAMETER SkopePath
    Where to load the modules. AllUsers = Default: Copy module to Program Files Directory
    CurrentUser = Copy module to user Documents\WindowsPowerShell Folder'
.PARAMETER CreateShortcut
    - Create desktop shortcut to load Powercli
.PARAMETER ForceInstall
    Force modules to re-import and install even if same version found
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -file "Install-PowerCLI.ps1" -CreateShortcut
.NOTES
    Script name: Install-PowerCLI.ps1
    Version:     3.1.0020
    Author:      Richard Tracy
    DateCreated: 2018-04-02
    LastUpdate:  2019-02-13

.LINK
    https://code.vmware.com/web/dp/tool/vmware-powercli/11.1.0
    http://www.powershellcrack.com/2017/09/installing-powercli-on-disconnected.html
    https://docs.microsoft.com/en-us/powershell/gallery/psget/repository/bootstrapping_nuget_proivder_and_exe
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Where to load the modules. AllUsers = Default: Copy module to Program Files Directory; 
                                                                                   CurrentUser = Copy module to user Documents\WindowsPowerShell Folder')]
	[ValidateSet("CurrentUser","AllUsers")]
    [string]$ScopePath = 'AllUsers',
    [Parameter(Mandatory=$false)]
    [switch]$CreateShortcut,
        [Parameter(Mandatory=$false,Position=1,HelpMessage='Force modules to re-import and install')]
	[switch]$ForceInstall = $false
)


##*===========================================================================
##* FUNCTIONS
##*===========================================================================
function Test-IsISE {
# try...catch accounts for:
# Set-StrictMode -Version latest
    try {    
        return $psISE -ne $null;
    }
    catch {
        return $false;
    }
}


Function Write-LogEntry {
    param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter(Mandatory=$false,Position=2)]
		[string]$Source = '',
        [parameter(Mandatory=$false)]
        [ValidateSet(0,1,2,3,4)]
        [int16]$Severity,

        [parameter(Mandatory=$false, HelpMessage="Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$OutputLogFile = $Global:LogFilePath,

        [parameter(Mandatory=$false)]
        [switch]$Outhost
    )
    ## Get the name of this function
    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

    [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
	[string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
	[int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes
	[string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias
    #  Get the file name of the source script

    Try {
	    If ($script:MyInvocation.Value.ScriptName) {
		    [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
	    }
	    Else {
		    [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
	    }
    }
    Catch {
	    $ScriptSource = ''
    }
    
    
    If(!$Severity){$Severity = 1}
    $LogFormat = "<![LOG[$Message]LOG]!>" + "<time=`"$LogTimePlusBias`" " + "date=`"$LogDate`" " + "component=`"$ScriptSource`" " + "context=`"$([Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + "type=`"$Severity`" " + "thread=`"$PID`" " + "file=`"$ScriptSource`">"
    
    # Add value to log file
    try {
        Out-File -InputObject $LogFormat -Append -NoClobber -Encoding Default -FilePath $OutputLogFile -ErrorAction Stop
    }
    catch {
        Write-Host ("[{0}] [{1}] :: Unable to append log entry to [{1}], error: {2}" -f $LogTimePlusBias,$ScriptSource,$OutputLogFile,$_.Exception.ErrorMessage) -ForegroundColor Red
    }
    If($Outhost){
        If($Source){
            $OutputMsg = ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$Source,$Message)
        }
        Else{
            $OutputMsg = ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$ScriptSource,$Message)
        }

        Switch($Severity){
            0       {Write-Host $OutputMsg -ForegroundColor Green}
            1       {Write-Host $OutputMsg -ForegroundColor Gray}
            2       {Write-Warning $OutputMsg}
            3       {Write-Host $OutputMsg -ForegroundColor Red}
            4       {If($Global:Verbose){Write-Verbose $OutputMsg}}
            default {Write-Host $OutputMsg}
        }
    }
}




Function Copy-ItemWithProgress
{
    [CmdletBinding()]
    Param
    (
    [Parameter(Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        Position=0)]
    $Source,
    [Parameter(Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        Position=0)]
    $Destination,
    [Parameter(Mandatory=$false,
        ValueFromPipelineByPropertyName=$true,
        Position=0)]
    [switch]$Force
    )

    $Source = $Source
    
    #get the entire folder structure
    $Filelist = Get-Childitem $Source -Recurse

    #get the count of all the objects
    $Total = $Filelist.count

    #establish a counter
    $Position = 0

    #Stepping through the list of files is quite simple in PowerShell by using a For loop
    foreach ($File in $Filelist)

    {
        #On each file, grab only the part that does not include the original source folder using replace
        $Filename = ($File.Fullname).replace($Source,'')
        
        #rebuild the path for the destination:
        $DestinationFile = ($Destination+$Filename)
        
        #show progress
        Write-Progress -Activity "Copying data from $source to $Destination" -Status "Copying File $Filename" -PercentComplete (($Position/$total)*100)
        
        #do copy (enforce
        Copy-Item $File.FullName -Destination $DestinationFile -Force:$Force -ErrorAction SilentlyContinue | Out-Null

        #bump up the counter
        $Position++
    }

}

##*===========================================================================
##* VARIABLES
##*===========================================================================
## Instead fo using $PSScriptRoot variable, use the custom InvocationInfo for ISE runs
## Variables: Script Name and Script Paths
[string]$scriptPath = $MyInvocation.MyCommand.Definition
#Since running script within Powershell ISE doesn't have a $scriptpath...hardcode it
If(Test-IsISE){$scriptPath = "C:\GitHub\PowerCLI-ModuleInstaller\Install-PowerCLI.ps1"}
[string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptPath)
[string]$scriptFileName = Split-Path -Path $scriptPath -Leaf
[string]$scriptDirectory = Split-Path -Path $scriptPath -Parent
[string]$invokingScript = (Get-Variable -Name 'MyInvocation').Value.ScriptName

Try
{
	$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $Progress = New-Object -ComObject Microsoft.SMS.TSprogressUI
	#$logPath = $tsenv.Value("LogPath")
    $LogPath = $tsenv.Value("_SMSTSLogPath")
}
Catch
{
	Write-Warning "TS environment not detected. Assuming stand-alone mode."
}

If(!$LogPath){$LogPath = $env:TEMP}
[string]$FileName = $scriptName +'.log'
$Global:LogFilePath = Join-Path $LogPath -ChildPath $FileName
Write-Host "Using log file: $LogFilePath"

#set variable to false
$InstallModule = $false
$UninstallModule = $false
[string]$originalLocation = Get-Location

#Get required folder and File paths
[string]$ModulesPath = Join-Path -Path $scriptDirectory -ChildPath 'Modules'
[string]$BinPath = Join-Path -Path $scriptDirectory -ChildPath 'Bin'
[string]$PSScriptsPath = Join-Path -Path $scriptDirectory -ChildPath 'Scripts'



#Get all paths to PowerShell Modules
$UserModulePath = $env:PSModulePath -split ';' | Where {$_ -like "$home*"}
$UserScriptPath = "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts"
$DefaultUserModulePath = "$env:SystemDrive\Users\Default\Documents\WindowsPowerShell\Modules"
$DefaultUserScriptPath = "$env:SystemDrive\Users\Default\Documents\WindowsPowerShell\Scripts"
$AllUsersModulePath = $env:PSModulePath -split ';' | Where {$_ -like "$env:ProgramFiles\WindowsPowerShell*"}
$AllUsersScriptsPath = ($env:PSModulePath -split ';' | Where {$_ -like "$env:ProgramFiles\WindowsPowerShell*"}) -replace "Modules", "Scripts"
$SystemModulePath = $env:PSModulePath -split ';' | Where {$_ -like "$env:windir*"}

#find profile module
$PowerShellNoISEProfile = $profile -replace "ISE",""


#Install Nuget prereq
$NuGetAssemblySourcePath = Get-ChildItem "$BinPath\nuget" -Recurse -Filter *.dll
If($NuGetAssemblySourcePath){
    $NuGetAssemblyVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($NuGetAssemblySourcePath.FullName).FileVersion
    $NuGetAssemblyDestPath = "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget\$NuGetAssemblyVersion"
    If (!(Test-Path $NuGetAssemblyDestPath)){
        Write-LogEntry ("Copying nuget Assembly [{0}] to [{1}]" -f $NuGetAssemblyVersion, $NuGetAssemblyDestPath) -Outhost
        New-Item $NuGetAssemblyDestPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path $NuGetAssemblySourcePath.FullName -Destination $NuGetAssemblyDestPath -Force -ErrorAction SilentlyContinue | Out-Null
    }
    #Unlock the dll incase it was downlaoded from the internt usng web browser
    Get-childitem $NuGetAssemblyDestPath -Filter *.dll | Unblock-File -Confirm:$false | Out-Null
}

##*===============================================
##* Load Section
##*===============================================
#Get-Module VMware.* -ListAvailable | Uninstall-Module -Force
$LatestVersion = $null
# Find Module name using regex
# Name can be anything as long as VMware and PowerCLI are near each other
$ModuleFolder = "\bVMware\b[^\n]+\bPowerCLI\b"

# Get Modules in modules folder and whats installed
#$Modules = Get-ChildItem -Path $ModulesPath -Recurse | Where-Object { $_.Name -match $query} | % {Get-ChildItem -Path $_.FullName -Filter *.psd1}
$ModuleSetSourcePath = Get-ChildItem -Path $ModulesPath | Where-Object { $_.PSIsContainer} | Where-Object { $_ -match $regex}

#if more than one module is found, try to find the latest
#$ModuleVersion = $Module.BaseName.split("-")[1]
#  this Used to detect by folder name, but this is safer
Foreach($Module in $ModuleSetSourcePath){
    $MainModuleFile = Get-ChildItem -Path $Module.FullName -Recurse -Filter *.psd1 | Where-Object { $_.Name -match $ModuleFolder}

    #grab manifest version
    $MainManifest = Test-ModuleManifest $MainModuleFile.FullName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    If($MainManifest.Name -match $ModuleFolder){
        If($MainManifest.Version -ge $LatestVersion){
            $LatestVersion = $MainManifest.Version
            $strLatestVersion = [string]$LatestVersion
            $ModulePath = $Module.FullName
            #isolate the module folder name
            $ModuleName = $MainManifest.Name
        }
    }
}

#if a module name was found continue
If($ModuleName){
    Write-LogEntry ("Downloaded module found [{0}] (version:{2}). Located in folder [{1}]" -f $ModuleName,$ModulePath,$strLatestVersion) -Outhost

    #Get destination path based on skopepath
    switch($ScopePath){
        "AllUsers"    { #$ModuleDestPath = Join-path $AllUsersModulePath -ChildPath $ModuleName
                        $ModuleDestPath = $AllUsersModulePath
                        #$ScriptsDestPath = Join-path $AllUsersScriptsPath -ChildPath $ModuleName
                        $ScriptsDestPath = $AllUsersScriptsPath
                      } 

        "CurrentUser" { #$ModuleDestPath = Join-path $UserModulePath -ChildPath $ModuleName
                        $ScriptsDestPath = $UserModulePath
                        #$ModuleDestPath = Join-path $UserModulePath -ChildPath $ModuleName
                        $ScriptsDestPath = $UserScriptPath
                      }

    }
}
Else{
    Write-LogEntry ("Found module [{0}]. The module does not match [VMware.PowerCLI]. Exiting..." -f $Module) -Severity 2 -Outhost
    Exit
}
##*===============================================
##* Detect Section
##*===============================================
#detect if Powercli is already installed
$FoundModule = Get-Module VMware* -ListAvailable
$FoundManifestModule = $FoundModule | Where {$_.ModuleType -eq "Manifest"}
$strFoundLatestVersion = [string]$FoundManifestModule.Version

If ($FoundModule){

    #check if installed version is the same as whats downloaded
    If($FoundModule.Version -ge $LatestVersion){
        
        #If forced, set trigger uninstall switch
        If($ForceInstall){
            Write-LogEntry ("[{0}] module is already installed. The version is equal to or greater than: {1} but the force switch is set to ignore..." -f $ModuleName,$strFoundLatestVersion) -Outhost
            $UninstallModule = $true
            $InstallModule = $true
        }
        Else{
            Write-LogEntry ("[{0}] module is already installed. The version is equal to or greater than: {1}" -f $ModuleName,$strFoundLatestVersion) -Outhost
        }
    }
    Else{
       $UninstallModule = $true
       $InstallModule = $true
    }
}
Else{
    Write-LogEntry ("[{0}] module is not found or installed" -f $ModuleName) -Outhost
    $InstallModule = $true
}

##*===============================================
##* Uninstall Section
##*===============================================
If($UninstallModule){
    #clear list, build collection object
    $Dependencies = @()
    $Dependencies = New-Object System.Collections.Generic.List[System.Object]

    # Grab root modules dependencies
    $Manifest = Test-ModuleManifest $FoundManifestModule.RootModule -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    $DependencyModules = $Manifest.RequiredModules
    
    #uninstall and remove each dependency
    Foreach ($Module in $DependencyModules){
        Try{
            Write-LogEntry ("Uninstalling module [{0} ({1})]..." -f $Module.Name,$Module.Version) -Outhost
            Get-Module $Module.Name | Remove-Module -Force
        }
        Catch{
            $ErrorMessage = $_.Exception.Message
            Write-LogEntry ("Failed to uninstall module [{0} ({1})]: {2}" -f $Module.Name,$Module.Version,$ErrorMessage) -Severity 3 -Outhost
            Break
        }
    }

    #uninstall and remove Main Module
    Try{
        Write-LogEntry ("Uninstalling module [{0} ({1})]..." -f $FoundModule.Name,$FoundModule.Version) -Outhost
        Get-Module $FoundManifestModule.Name | Remove-Module -Force
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        Write-LogEntry ("Failed to uninstall module [{0} ({1})]: {2}" -f $FoundModule.Name,$FoundModule.Version,$ErrorMessage) -Severity 3 -Outhost
        break 
    }
}

##*===============================================
##* Install Section
##*===============================================
#install new module if force or not installed already
If($ForceInstall -or $InstallModule){
    #must be copied to one of PowerShell module folders
    Write-LogEntry ("Copying VMWare PowerCLI ({0}) to [{1}]" -f $strLatestVersion,$ModuleDestPath) -Severity 1 -Outhost
    Copy-ItemWithProgress -Source $ModulePath -Destination $ModuleDestPath -Force

    #set the location of the destination folder to install it
    Set-Location -Path $ModuleDestPath | Out-Null

    #install the modules
    #this will copy it to the root module folder with versions folders for each dependency
    Try{
        Write-LogEntry ("Installing VMWare.PowerCLI ({0}) to [{1}]" -f $strLatestVersion,$ModuleDestPath) -Severity 1 -Outhost
        Get-Module VMware.PowerCLI* -ListAvailable -Refresh | Import-Module
        #Install-Module $ModuleName -Scope $ScopePath -AllowClobber -Force -Confirm:$False -ErrorAction Stop -ErrorVariable err | Out-Null
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        Write-LogEntry ("Failed to Install VMWare.PowerCLI Module ({0}): {1}" -f $strLatestVersion,$ErrorMessage) -Severity 3 -Outhost
        Exit -1 

    }

    #Change location so it not in use
    Set-Location -Path $originalLocation
}

##*===============================================
##* Shortcut Section
##*===============================================
If($CreateShortcut){
    #copy PowerCLI Modules to User directory if they don't exist ($env:PSModulePath)
    Write-LogEntry ("Checking installed module: [{0}\{1}\{2}.psd1]" -f $ModuleDestPath,$strLatestVersion,$ModuleName) -Severity 4 -Outhost

    If (Test-Path "$AllUsersModuleDestPath\$ModuleFolder\$ModuleVersion\$ModuleFolder.psd1"){

        #copy modules files
        Write-LogEntry ("{0} :: Copying VMware PowerCLI: [{1}] modules to [{2}]" -f $scriptName,$ModuleVersion,$ModuleDestPath) -Severity 1 -Outhost
        Copy-Item -Path "$PSScriptsPath\$ModuleFolder\VMware.PowerCLI.ico" -Destination $ModuleDestPath -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path "$PSScriptsPath\$ModuleFolder\Initialize-PowerCLIEnvironment.ps1" -Destination $ModuleDestPath -ErrorAction SilentlyContinue | Out-Null

        #copy scripts for module files for public desktop
        Write-LogEntry ("{0} :: Copying VMware PowerCLI shortcuts to: [{1}\Desktop]" -f $scriptName,$env:PUBLIC) -Severity 1 -Outhost
        Copy-Item -Path "$PSScriptsPath\$ModuleFolder\VMware PowerCLI (32-Bit).lnk" -Destination "$env:PUBLIC\Desktop" -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path "$PSScriptsPath\$ModuleFolder\VMware PowerCLI.lnk" -Destination "$env:PUBLIC\Desktop" -ErrorAction SilentlyContinue | Out-Null
    }
}
