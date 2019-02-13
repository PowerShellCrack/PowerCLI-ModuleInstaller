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
    Version:     3.0
    Author:      Richard Tracy
    DateCreated: 2018-04-02
    LastUpdate:  2019-02-12

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
    [string]$SkopePath = 'AllUsers',
    [Parameter(Mandatory=$false)]
    [switch]$CreateShortcut,
        [Parameter(Mandatory=$false,Position=1,HelpMessage='Force modules to re-import and install')]
	[switch]$ForceInstall = $false
)


##*===========================================================================
##* FUNCTIONS
##*===========================================================================
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


function Copy-WithProgress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Source,
        [Parameter(Mandatory = $true)]
        [string] $Destination,
        [int] $Gap = 0,
        [int] $ReportGap = 200,
        [ValidateSet("Directories","Files")]
        [string] $ExcludeType,
        [string] $Exclude,
        [string] $ProgressDisplayName
    )
    # Define regular expression that will gather number of bytes copied
    $RegexBytes = '(?<=\s+)\d+(?=\s+)';

    #region Robocopy params
    # MIR = Mirror mode
    # NP  = Don't show progress percentage in log
    # NC  = Don't log file classes (existing, new file, etc.)
    # BYTES = Show file sizes in bytes
    # NJH = Do not display robocopy job header (JH)
    # NJS = Do not display robocopy job summary (JS)
    # TEE = Display log in stdout AND in target log file
    # XF file [file]... :: eXclude Files matching given names/paths/wildcards.
    # XD dirs [dirs]... :: eXclude Directories matching given names/paths.
    $CommonRobocopyParams = '/MIR /NP /NDL /NC /BYTES /NJH /NJS';
    
    switch ($ExcludeType){
        Files { $CommonRobocopyParams += ' /XF {0}' -f $Exclude };
	    Directories { $CommonRobocopyParams += ' /XD {0}' -f $Exclude };
    }
    
    #endregion Robocopy params
    
    #generate log format
    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'

    #region Robocopy Staging
    Write-Verbose -Message 'Analyzing robocopy job ...';
    $StagingLogPath = '{0}\offlinemodules-staging-{1}.log' -f $env:temp, (Get-Date -Format 'yyyy-MM-dd hh-mm-ss');

    $StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams;
    Write-Verbose -Message ('Staging arguments: {0}' -f $StagingArgumentList);
    Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList -WindowStyle Hidden;
    # Get the total number of files that will be copied
    $StagingContent = Get-Content -Path $StagingLogPath;
    $TotalFileCount = $StagingContent.Count - 1;

    # Get the total number of bytes to be copied
    [RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | % { $BytesTotal = 0; } { $BytesTotal += $_.Value; };
    Write-Verbose -Message ('Total bytes to be copied: {0}' -f $BytesTotal);
    #endregion Robocopy Staging

    #region Start Robocopy
    # Begin the robocopy process
    $RobocopyLogPath = '{0}\offlinemodules-{1}.log' -f $env:temp, (Get-Date -Format 'yyyy-MM-dd hh-mm-ss');
    $ArgumentList = '"{0}" "{1}" /LOG:"{2}" /ipg:{3} {4}' -f $Source, $Destination, $RobocopyLogPath, $Gap, $CommonRobocopyParams;
    Write-Verbose -Message ('Beginning the robocopy process with arguments: {0}' -f $ArgumentList);
    $Robocopy = Start-Process -FilePath robocopy.exe -ArgumentList $ArgumentList -Verbose -PassThru -WindowStyle Hidden;
    Start-Sleep -Milliseconds 100;
    #endregion Start Robocopy

    #region Progress bar loop
    while (!$Robocopy.HasExited) {
        Start-Sleep -Milliseconds $ReportGap;
        $BytesCopied = 0;
        $LogContent = Get-Content -Path $RobocopyLogPath;
        $BytesCopied = [Regex]::Matches($LogContent, $RegexBytes) | ForEach-Object -Process { $BytesCopied += $_.Value; } -End { $BytesCopied; };
        $CopiedFileCount = $LogContent.Count - 1;
        Write-Verbose -Message ('Bytes copied: {0}' -f $BytesCopied);
        Write-Verbose -Message ('Files copied: {0}' -f $LogContent.Count);
        $Percentage = 0;
        if ($BytesCopied -gt 0) {
           $Percentage = (($BytesCopied/$BytesTotal)*100)
        }
        If ($ProgressDisplayName){$ActivityDisplayName = $ProgressDisplayName}Else{$ActivityDisplayName = 'Robocopy'}
        Write-Progress -Activity $ActivityDisplayName -Status ("Copied {0} of {1} files; Copied {2} of {3} bytes" -f $CopiedFileCount, $TotalFileCount, $BytesCopied, $BytesTotal) -PercentComplete $Percentage
    }
    #endregion Progress loop

    #region Function output
    [PSCustomObject]@{
        BytesCopied = $BytesCopied;
        FilesCopied = $CopiedFileCount;
    };
    #endregion Function output
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
[string]$FileName = $scriptRoot +'.log'
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
    Write-LogEntry ("Found lastest [{0}] module in folder [{1}] with version: {2}" -f $ModuleName,$ModulePath,$strLatestVersion) -Outhost

    #Get destination path based on skopepath
    switch($SkopePath){
        "AllUsers"    { $ModuleDestPath = Join-path $AllUsersModulePath -ChildPath $ModuleName
                        $ScriptsDestPath = Join-path $AllUsersScriptsPath -ChildPath $ModuleName
                      } 

        "CurrentUser" { $ModuleDestPath = Join-path $UserModulePath -ChildPath $ModuleName
                        $ScriptsDestPath = Join-path $UserScriptPath -ChildPath $ModuleName
                      }

    }
}
Else{
    Write-LogEntry ("module name [{0}] does not match VMware.PowerCLI" -f $Module) -Severity 2 -Outhost
    Exit
}
##*===============================================
##* Detect Section
##*===============================================
#detect if Powercli is already installed
$FoundModule = Get-Module VMware.PowerCLI* -ListAvailable

If ($FoundModule){

    #check if installed version is the same as whats downloaded
    If($FoundModule.Version -ge $LatestVersion){
        
        #If forced, set trigger uninstall switch
        If($ForceInstall){
            Write-LogEntry ("An already installed [{0}] module version is equal to or greater than: {1}, Force switch is ignoring version..." -f $ModuleName,$strLatestVersion) -Outhost
            $UninstallModule = $true
        }
        Else{
            Write-LogEntry ("An already installed [{0}] module version is equal to or greater than: {1}" -f $ModuleName,$strLatestVersion) -Outhost
        }
    }
    Else{
       $UninstallModule = $true
       $InstallModule = $true
    }
}
Else{
    Write-LogEntry ("[{0}] module is not installed. Preparing to install..." -f $ModuleName) -Outhost
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
    $Manifest = Test-ModuleManifest $InstalledPowerCLI.RootModule -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    $DependencyModules = $Manifest.RequiredModules
    
    #uninstall and remove each dependency
    Foreach ($Module in $DependencyModules){
        Try{
            Write-LogEntry ("Uninstalling module: {0} ({1})..." -f $Module.Name,$Module.Version) -Outhost
            Uninstall-Module -Name $Module.Name -MinimumVersion $Module.Version -Force -ErrorAction Stop | Remove-Module -Force
        }
        Catch{
            Write-LogEntry ("Failed to uninstall module: {0} ({1})" -f $Module.Name,$Module.Version) -Severity 3 -Outhost
            break 
        }
    }

    #uninstall and remove Main Module
    Try{
        Write-LogEntry ("Uninstalling module: {0} ({1})..." -f $FoundModule.Name,$FoundModule.Version) -Outhost
        Uninstall-Module -Name $FoundModule.Name -MinimumVersion $FoundModule.Version -Force -ErrorAction Stop | Remove-Module -Force
    }
    Catch{
        Write-LogEntry ("Failed to uninstall module: {0} ({1})" -f $FoundModule.Name,$FoundModule.Version) -Severity 3 -Outhost
        break 
    }
}

##*===============================================
##* Install Section
##*===============================================
#install new module if force or not installed already
If($ForceInstall -or $InstallModule){
    #copy Module to the STAGING folder
    $StagingModuleDestPath = ($ModuleDestPath + '-stage')
    #must be copied to one of PowerShell module folders
    Write-LogEntry ("Installing VMWare PowerCLI ({0}) to [{1}]" -f $strLatestVersion,$ModuleDestPath) -Severity 1 -Outhost
    Copy-WithProgress -Source $ModulePath -Destination $StagingModuleDestPath -ExcludeType Directories -Exclude 'nuget' -ProgressDisplayName ('Copying {0} ({1}) Modules Files...' -f $ModuleName,$strLatestVersion)

    #set the location of the destination folder to install it
    Set-Location -Path $StagingModuleDestPath

    #install the modules
    #this will copy it to the root module folder with versions folders for each dependency
    Try{
        Write-LogEntry ("Installing VMWare PowerCLI ({0}) to [{1}]" -f $strLatestVersion,$ModuleDestPath) -Severity 1 -Outhost
        Install-Module $ModuleName -Scope $SkopePath -AllowClobber -Force -Confirm:$False | Out-Null
    }
    Catch{
        Write-LogEntry ("Failed to Install VMWare PowerCLI Module ({0})..." -f $strLatestVersion) -Severity 3 -Outhost
        Exit -1 

    }

    #Change location so it not in use
    Set-Location -Path $originalLocation

    #delete original staging folder
    Remove-Item $StagingModuleDestPath -Recurse -Force | Out-Null
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
    Exit
}
Else{
    Import-Module $ModuleName
}