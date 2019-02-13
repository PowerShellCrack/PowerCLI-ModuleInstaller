<#
.SYNOPSIS
    Saves PowerShell Modules for import to diconnect networks
.DESCRIPTION
    Run this script on a internet connected system
    this script will download latest nuget assembly with packagemanagement modules
    plus any additional module found. Required for disconnected system
.PARAMETER Install
    Install modules on online system as well
.PARAMETER RemoveOld
    Remove older modules if found
.PARAMETER ForceInstall
    Force modules to re-import and install if same version found
.PARAMETER Refresh
    Re-Download modules if exist
.NOTES
    https://docs.microsoft.com/en-us/powershell/gallery/psget/repository/bootstrapping_nuget_proivder_and_exe
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Remove older modules if found')]
	[switch]$RemoveOld = $true,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Install modules on online system as well')]
	[switch]$Install = $true,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Force modules to re-import and install')]
	[switch]$ForceInstall = $false,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Re-Download modules if exist')]
	[switch]$Refresh = $false
)
##*===============================================
##* VARIABLE DECLARATION
##*===============================================
## Variables: Script Name and Script Paths
[string]$scriptPath = $MyInvocation.MyCommand.Definition
#Since running script within Powershell ISE doesn't have a $scriptpath...hardcode it
If(Test-IsISE){$scriptPath = "C:\GitHub\PowerCLI-ModuleInstaller\Get-PowerCli.ps1"}
[string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptPath)
[string]$scriptFileName = Split-Path -Path $scriptPath -Leaf
[string]$scriptRoot = Split-Path -Path $scriptPath -Parent
[string]$invokingScript = (Get-Variable -Name 'MyInvocation').Value.ScriptName

#Get required folder and File paths
[string]$ModulesPath = Join-Path -Path $scriptRoot -ChildPath 'Modules'
[string]$BinPath = Join-Path -Path $scriptDirectory -ChildPath 'Bin'

$OnlineModules = "VMware.PowerCLI"

##*===============================================
##*==========================================================================
##* FUNCTION
##*==========================================================================
Function Output-Prefix{
    [string]$timezone = [Regex]::Replace([System.TimeZoneInfo]::Local.StandardName, '([A-Z])\w+\s*', '$1')
    [string]$date = Get-Date -Format ("ddd MMM dd hh:mm:ss {0} yyyy" -f $timezone) 
    return $date
}
$Prefix = Output-Prefix

#See if system is conencted to the internet
$internetConnected = Test-NetConnection www.powershellgallery.com -CommonTCPPort HTTP -InformationLevel Quiet -WarningAction SilentlyContinue

If($internetConnected)
{
    $Nuget = Install-PackageProvider Nuget –force –verbose
    $NuGetAssemblyVersion = $($Nuget).version
    Write-Host "INSTALLED: Nuget [$NuGetAssemblyVersion] was installed" -ForegroundColor Green
    #Copy Nuget prereqs
    $NuGetAssemblyDestPath = Get-ChildItem "$BinPath\nuget" -Filter *.dll -Recurse
    $NuGetAssemblySourcePath = "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget"
    If ($NuGetAssemblyDestPath.FullName)
    {
        If($Refresh){
            Write-Host "BACKUP: Copying nuget Assembly [$NuGetAssemblyVersion] from $NuGetAssemblySourcePath" -ForegroundColor Gray
            Copy-Item "$NuGetAssemblySourcePath\$NuGetAssemblyVersion\Microsoft.PackageManagement.NuGetProvider.dll" $NuGetAssemblyDestPath.FullName -Force -ErrorAction SilentlyContinue
        }
        Else{
            Write-Host "FOUND: Nuget [$NuGetAssemblyVersion] already copied" -ForegroundColor Green
        }
    }
    Else{
        Write-Host "BACKUP: Copying nuget Assembly [$NuGetAssemblyVersion] from $NuGetAssemblySourcePath" -ForegroundColor Gray
        New-Item "$ModulesPath\nuget\$NuGetAssemblyVersion" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Copy-Item "$NuGetAssemblySourcePath\$NuGetAssemblyVersion\Microsoft.PackageManagement.NuGetProvider.dll" "$BinPath\nuget\$NuGetAssemblyVersion" -ErrorAction SilentlyContinue
    }

    #loop through each module
    Foreach ($Module in $OnlineModules){
        
        #get the module if found online
        $ModuleFound = Find-Module $Module
        If($ModuleFound)
        {
            [string]$ModuleVersion = $ModuleFound.Version
            [string]$ModuleName = $ModuleFound.Name
            
            #If specified, remove older modules in DownloadedModule directory if found
            If($RemoveOld)
            {
                $LikeModulesExist = Get-ChildItem $ModulesPath -Directory | Where-Object {$_.FullName -match "$ModuleName" -and $_.FullName -notmatch "$ModuleName-$ModuleVersion"} | foreach {
                        $_ | Remove-Item -Force -Recurse
                        Write-host "REMOVED: $($_.FullName)" -ForegroundColor DarkYellow
                    }
            }


            #Check to see it module is already downloaded
            If(Test-Path "$ModulesPath\$ModuleName-$ModuleVersion")
            {
                #If specified, Re-Download modules 
                If($Refresh)
                {
                    Write-Host "BACKUP: $ModuleName [$ModuleVersion] found but will be re-downloaded..." -ForegroundColor Yellow
                    Save-Module -Name $ModuleName -Path $ModulesPath\$ModuleName-$ModuleVersion -Force
                }
                Else{
                    Write-Host "FOUND: $ModuleName [$ModuleVersion] already downloaded" -ForegroundColor Gray
                }
            }
            Else{
                Write-Host "BACKUP: $ModuleName [$ModuleVersion] not found, downloading for offline install" -ForegroundColor Gray
                New-Item "$ModulesPath\$ModuleName-$ModuleVersion" -ItemType Directory | Out-Null
                Save-Module -Name $ModuleName -Path $ModulesPath\$ModuleName-$ModuleVersion
            }

            #If specified, Install modules on local system as well 
            If($Install)
            {
                If([string](Get-Module $Module).Version -ne $ModuleVersion -or $ForceInstall){
                    Try{
                        Write-Host "INSTALL: $Module [$ModuleVersion] will be installed locally as well, please wait..." -ForegroundColor Gray
                        Install-Module $Module -AllowClobber -SkipPublisherCheck -Force
                        Import-Module $Module
                    }
                    Catch{
				        Write-Output ("[{0}][{1}] Failed to install and import  $Module" -f " $Module",$Prefix)
				        Write-Output $_.Exception | format-list -force
				        Exit $_.ExitCode
			        }
                }
                Else{
                    Write-Host "INSTALL: $Module [$ModuleVersion] is already installed, skipping install..." -ForegroundColor Gray
                }
            }
        }
        Else{
            Write-Host "WARNING: $Module was not found online" -ForegroundColor Yellow
        }

        Write-Host "COMPLETED: Done working with module: $Module" -ForegroundColor Green

    } #End Loop
}
Else{
    Write-Host "ERROR: Unable to connect to the internet to grab modules" -ForegroundColor Red
    throw $_.error
}
