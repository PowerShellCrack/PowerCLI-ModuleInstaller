<#
.SYNOPSIS
    Install PowerCLI Offline
.DESCRIPTION
    Install the nuget package manamgment prereq then installs PowerCLI module
.PARAMETER 
    NONE
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -file "Install-PowerCLI.ps1"
.NOTES
    Script name: Install-PowerCLI.ps1
    Version:     2.0
    Author:      Richard Tracy
    DateCreated: 2018-04-02
    LastUpdate:  2019-02-08

.LINK
    https://code.vmware.com/web/dp/tool/vmware-powercli/11.1.0
    http://www.powershellcrack.com/2017/09/installing-powercli-on-disconnected.html
#>


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



##*===========================================================================
##* VARIABLES
##*===========================================================================
## Instead fo using $PSScriptRoot variable, use the custom InvocationInfo for ISE runs
If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
[string]$scriptDirectory = Split-Path $MyInvocation.MyCommand.Path -Parent
[string]$scriptName = Split-Path $MyInvocation.MyCommand.Path -Leaf
[string]$scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($scriptName)

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
[string]$FileName = $scriptBaseName +'.log'
$Global:LogFilePath = Join-Path $LogPath -ChildPath $FileName
Write-Host "Using log file: $LogFilePath"


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
$SystemModulePath = $env:PSModulePath -split ';' | Where {$_ -like "$env:windir*"}

#find profile module
$PowerShellNoISEProfile = $profile -replace "ISE",""


#Install Nuget prereq
$NuGetAssemblySourcePath = Get-ChildItem "$BinPath\nuget" -Recurse -Filter *.dll
If($NuGetAssemblySourcePath){
    $NuGetAssemblyVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($NuGetAssemblySourcePath.FullName).FileVersion
    $NuGetAssemblyDestPath = "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget\$NuGetAssemblyVersion"
    If (!(Test-Path $NuGetAssemblyDestPath)){
        Write-Host "Copying nuget Assembly ($NuGetAssemblyVersion) to $NuGetAssemblyDestPath" -ForegroundColor Cyan
        New-Item $NuGetAssemblyDestPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path $NuGetAssemblySourcePath.FullName -Destination $NuGetAssemblyDestPath -Force -ErrorAction SilentlyContinue | Out-Null
    }
}


#Find PowerCLI Module
$ModuleFolder = $("VMware.PowerCLI")


# Get Modules in modules folder and whats installed
$ModuleSetSourcePath = Get-ChildItem -Path $ModulesPath -Depth 0 | Where-Object { $_.PSIsContainer} | Where-Object { $_ -like "$ModuleFolder*"}
$ModuleVersion = $ModuleSetSourcePath.BaseName.split("-")[1]



$UserModuleDestPath = "$UserModulePath\$ModuleFolder"
$DefaultUserModuleDestPath = "$DefaultUserModulePath\$ModuleFolder"
$AllUsersModuleDestPath = "$AllUsersModulePath\$ModuleFolder"


Set-Location -Path $ModuleSetSourcePath.FullName
Write-LogEntry ("Installing VMWare PowerCLI ({0}) to [{1}]" -f $ModuleVersion,$AllUsersModuleDestPath) -Severity 1 -Outhost
Try{
    Install-Module $ModuleFolder -Scope AllUsers -AllowClobber -Force -Confirm:$False | Out-Null
}
Catch{
    Write-LogEntry "Failed to Install VMWare PowerCLI ({0}) Module..." -Severity 3 -Outhost
    Exit -1 

}

#copy PowerCLI Modules to User directory if they don't exist ($env:PSModulePath)
Write-LogEntry ("Checking installed module: [{0}\{1}\{2}.psd1]" -f $AllUsersModuleDestPath,$ModuleVersion,$ModuleFolder) -Severity 4 -Outhost

If (Test-Path "$AllUsersModuleDestPath\$ModuleFolder\$ModuleVersion\$ModuleFolder.psd1"){

    #copy modules files
    Write-LogEntry ("{0} :: Copying VMware PowerCLI: [{1}] modules to [{2}]" -f $scriptName,$ModuleVersion,$AllUsersModuleDestPath) -Severity 1 -Outhost
    Copy-Item -Path "$PSScriptsPath\$ModuleFolder\VMware.PowerCLI.ico" -Destination $AllUsersModuleDestPath -ErrorAction SilentlyContinue | Out-Null
    Copy-Item -Path "$PSScriptsPath\$ModuleFolder\Initialize-PowerCLIEnvironment.ps1" -Destination $AllUsersModuleDestPath -ErrorAction SilentlyContinue | Out-Null

    #copy scripts for module files for public desktop
    Write-LogEntry ("{0} :: Copying VMware PowerCLI shortcuts to: [{1}\Desktop]" -f $scriptName,$env:PUBLIC) -Severity 1 -Outhost
    Copy-Item -Path "$PSScriptsPath\$ModuleFolder\VMware PowerCLI (32-Bit).lnk" -Destination "$env:PUBLIC\Desktop" -ErrorAction SilentlyContinue | Out-Null
    Copy-Item -Path "$PSScriptsPath\$ModuleFolder\VMware PowerCLI.lnk" -Destination "$env:PUBLIC\Desktop" -ErrorAction SilentlyContinue | Out-Null
}

Exit