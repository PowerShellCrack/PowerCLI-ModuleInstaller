<#
.SYNOPSIS
    Saves PowerShell Modules for import to disconnect networks
.DESCRIPTION
    Run this script on a internet connected system
    this script will download latest nuget assembly with packagemanagement modules
    plus any additional module found. Required for disconnected system
.PARAMETER Install
    Install modules on online system as well
.PARAMETER RemoveOld
    Remove older modules if found
.PARAMETER ForceInstall
    Force modules to re-import and install even if same version found
.PARAMETER Refresh
    Re-Download modules if exist
.NOTES
    Script name: Get-PowerCLI.ps1
    Version:     3.1.0020
    Author:      Richard Tracy
    DateCreated: 2018-04-02
    LastUpdate:  2019-02-13
.LINKS
    https://docs.microsoft.com/en-us/powershell/gallery/psget/repository/bootstrapping_nuget_proivder_and_exe
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false,Position=0,HelpMessage='Specify modules to download. for multiple, separate by commas')]
	[string[]]$OnlineModules = "VMware.PowerCLI",
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Remove older modules if found')]
	[switch]$RemoveOld = $true,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Install modules on online system as well')]
	[switch]$InstallAsWell = $false,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Force modules to re-import and install')]
	[switch]$ForceInstall = $false,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Re-Download modules if exist')]
	[switch]$Refresh = $false,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Does an extensive search on already downloaded modules and tries to name them using module standarization')]
	[switch]$FixInvalidModule = $false
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

Function Copy-ItemWithProgress
{
    [CmdletBinding()]
    Param
    (
    [Parameter(Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        Position=0)]
    [string]$Source,
    [Parameter(Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        Position=1)]
    [string]$Destination,
    [Parameter(Mandatory=$false,
        ValueFromPipelineByPropertyName=$true,
        Position=2)]
    [int16]$ParentID,
    [Parameter(Mandatory=$false,
        ValueFromPipelineByPropertyName=$true,
        Position=3)]
    [switch]$Force
    )

    Begin{
        $Source = $Source
    
        #get the entire folder structure
        $Filelist = Get-Childitem $Source -Recurse

        #get the count of all the objects
        $Total = $Filelist.count

        #establish a counter
        $Position = 0

        #set an id for the progress bar
        If($ParentID){$ParentID = $ParentID;$ThisProgressID = ($ParentID+1)}Else{$ProgressID = 1;$ThisProgressID = 1}
    }
    Process{
        #Stepping through the list of files is quite simple in PowerShell by using a For loop
        foreach ($File in $Filelist)

        {
            #On each file, grab only the part that does not include the original source folder using replace
            $Filename = ($File.Fullname).replace($Source,'')
        
            #rebuild the path for the destination:
            $DestinationFile = ($Destination+$Filename)
        
            #get just the folder path
            $DestinationPath = Split-Path $DestinationFile -Parent

            #show progress
            Write-Progress -Activity "Copying data from $source to $Destination" -Status "Copying File $Filename" -PercentComplete (($Position/$total)*100) -Id $ThisProgressID -ParentId $ParentID
        
            #create destination directories
            If (-not (Test-Path $DestinationPath) ) {
                New-Item $DestinationPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
            }

            #do copy (enforce)
            Try{
                Copy-Item $File.FullName -Destination $DestinationFile -Force:$Force -ErrorAction:$VerbosePreference -Verbose:($PSBoundParameters['Verbose'] -eq $true) | Out-Null
                Write-Verbose ("Copied file [{0}] to [{1}]" -f $File.FullName,$DestinationFile)
            }
            Catch{
                Write-Host ("Unable to copy file in {0} to {1}; Error: {2}" -f $File.FullName,$DestinationFile ,$_.Exception.Message) -ForegroundColor Red
                break
            }
            #bump up the counter
            $Position++
        }
    }
    End{}
}


Function Inventory-ModuleStructure{
    [CmdletBinding(DefaultParametersetName="none")] 
    Param
    (
    [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
    [ValidateScript({Test-Path $_ -PathType 'Container'})] 
    [String]$ModulePath,
    
    [Parameter(ParameterSetName="Validate",Mandatory=$false)]
    [switch]$Validate,

    [Parameter(ParameterSetName="Fix",Mandatory=$true)]
    [switch]$Fix,

    [Parameter(ParameterSetName="Fix",Mandatory=$false)]
    [ValidateScript({Test-Path $_ -PathType 'Container'})]
    [String]$DestRootPath,

    [Parameter(ParameterSetName="Fix",Mandatory=$false)]
    [switch]$Overwrite

    )
    Begin{
        $ParamSetName = $PsCmdLet.ParameterSetName
        Write-verbose ('Set name is: {0}' -f $ParamSetName)
        
        #grab base directory for use later
        $BasePath = Split-Path $ModulePath -Parent

        $rootFolders = Get-ChildItem $ModulePath | where {$_.PSisContainer -eq $true}


        #get the count of all the objects
        $Total = $rootFolders.count

        #establish a counter
        $Position = 0

        $moduleCollection = @()
    }
    Process{
        
        #first get the main module manifest
        #grab psd1 module file in the module folder (ignore others that may exist in sub folders)
        #look to see if RootModule is not specified in the psd1 (these are considered manifest modules)
        Foreach ($folder in $rootFolders) {
            $Psd1 = Get-ChildItem $folder.FullName -Filter *.psd1 -Recurse | Select -First 1
            $Psd1Folder = Split-Path $Psd1.FullName -Parent 
            $Manifest = Test-ModuleManifest $Psd1.FullName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            Write-Verbose ("`$folder = {0}" -f $folder.FullName)
            If(-not($Manifest.RootModule)){   
                
                $MainModuleFolder = $Psd1Folder
                $MainModuleName = $Manifest.Name
                $MainModuleVersion = $Manifest.Version

                Write-Verbose ("`$MainModuleFolder = {0}" -f $MainModuleFolder)
                Write-Verbose ("`$MainModuleName = {0}" -f $MainModuleName)
                Write-Verbose ("`$MainModuleVersion = {0}" -f $MainModuleVersion)
                break
            }
            
        }


        #Now process all folders again and grab everything
        Foreach ($folder in $rootFolders) {
            Write-Verbose ("============================")      
            Write-Verbose ("`$folder = {0}" -f $folder.FullName)

            #grab psd1 module file in the module folder (ignore others that may exist in sub folders)
            $Psd1 = Get-ChildItem $folder.FullName -Filter *.psd1 -Recurse | Select -First 1
            Write-Verbose ("`$Psd1 = {0}" -f $Psd1)

            $Psd1Folder = Split-Path $Psd1.FullName -Parent
            Write-Verbose ("`$Psd1Folder = {0}" -f $Psd1Folder)

            #this will grab just the root path of the module
            $RootPath = Split-Path $folder.FullName -Parent
            Write-Verbose ("`$RootPath  = {0}" -f $RootPath)

            #get the modules folder structure by removing the root module path
            $CurrentFolder = $Psd1Folder.replace($BasePath + "\","")
            Write-Verbose ("`$CurrentFolder = {0}" -f $CurrentFolder)

            #parse the manifest file for proper name and version
            $Manifest = Test-ModuleManifest $Psd1.FullName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            
            Write-Verbose ("`$ManifestName = {0}" -f $Manifest.Name)
            Write-Verbose ("`$ManifestVersion = {0}" -f $Manifest.Version)

            #build new path based on standard naming: <modulefolder>-<moduleversion>\<modulefolder>\<moduleversion>\<subfolder>\<file>
            $NewFolder = ($MainModuleName + "-" + $MainModuleVersion + "\" + $Manifest.Name + "\" + $Manifest.Version)
                
            Write-Verbose ("`$NewFolder = {0}" -f $NewFolder)

            If($Psd1Folder -eq $MainModuleFolder){$ModuleType = 'Manifest'}Else{$ModuleType = 'Script'}

            #build module object
            $ModuleFileObject = New-Object PSObject -Property @{
                CurrentFolder = $CurrentFolder
                NewFolder = $NewFolder
                Manifest = $Psd1.name
                Module = $Manifest.RootModule
                Version = $Manifest.Version
                ModuleType = $ModuleType 
                RootModule = $MainModuleName
                RootModuleVersion = $MainModuleVersion
            }

            #collect objects
            $moduleCollection += $ModuleFileObject
        }
    }
    End{
        If($ParamSetName -eq "Fix"){
            Foreach ($moduleItem in $moduleCollection) {
                Write-Progress -Activity ("Copying files for module [{0}]" -f $moduleItem.Module) -Status ("Copying contents in folder [{0}]" -f $moduleItem.CurrentFolder) -PercentComplete (($Position/$total)*100) -ParentId 1

                #if destination path is not specified, then use source module path
                If(!$DestRootPath){$DestRootPath = Split-Path $ModulePath -Parent}

                #source folder
                $SourcePath = $BasePath + "\" + $moduleItem.CurrentFolder
                
                #include a root path with the new folder
                $DestPath = $DestRootPath + "\" + $moduleItem.NewFolder
                 
                Try{
                    Copy-ItemWithProgress $SourcePath $DestPath -ParentID 1 -Force:$Overwrite -Verbose:($PSBoundParameters['Verbose'] -eq $true)
                    Write-Verbose ("Copied files in [{0}] to [{1}]" -f $SourcePath,$DestPath)
                    $FixedDir = $true
                }
                Catch{
                    Write-Host ("Unable to copy file in {0} to {1}; Error: {2}" -f $SourcePath,$DestPath,$_.Exception.Message) -ForegroundColor Red
                    $FixedDir = $false
                    break
                }

                #bump up the counter
                $Position++
            }
        }

        #if fix not specified or comes back with errors
        If($ParamSetName -eq "Validate"){
            $modulevalidated=$true
            #compare the current folder with new folder on each module
            #if they are the same the module has a valid structure
            #if not, either use the fix switch or re-download
            Foreach ($moduleItem in $moduleCollection) {
                If($moduleItem.CurrentFolder -ne $moduleItem.NewFolder){
                    $modulevalidated=$false
                } 
            }
            return $modulevalidated  
        }
        ElseIf($FixedDir){
            $FixFolder = ($DestRootPath + "\" + $MainModuleName + "-" + $MainModuleVersion)
            return $FixFolder
        }
        Else{
            return $moduleCollection
        }

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


##*===============================================
##* VARIABLE DECLARATION
##*===============================================
## Variables: Script Name and Script Paths
[string]$scriptPath = $MyInvocation.MyCommand.Definition
#Since running script within Powershell ISE doesn't have a $scriptpath...hardcode it
If(Test-IsISE){$scriptPath = "D:\Development\GitHub\PowerCLI-ModuleInstaller\Get-PowerCLI.ps1"}
[string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptPath)
[string]$scriptFileName = Split-Path -Path $scriptPath -Leaf
[string]$scriptRoot = Split-Path -Path $scriptPath -Parent
[string]$invokingScript = (Get-Variable -Name 'MyInvocation').Value.ScriptName

#Get required folder and File paths
[string]$ModulesPath = Join-Path -Path $scriptRoot -ChildPath 'Modules'
[string]$BinPath = Join-Path -Path $scriptRoot -ChildPath 'Bin'




##*===============================================
##* Nuget Section
##*===============================================
#See if system is conencted to the internet
$internetConnected = Test-NetConnection www.powershellgallery.com -CommonTCPPort HTTP -InformationLevel Quiet -WarningAction SilentlyContinue | Out-NUll

If($internetConnected)
{
    #package management is a requirement for PowerShell modules
    $PackageManagement = Install-PackageProvider Nuget –force
    $PackageManagementAssemblyVersion = $($PackageManagement).version
    Write-Host ("Nuget Package Manger [{0}] is installed" -f $PackageManagementAssemblyVersion) -ForegroundColor Green

    #nuget is an open-source module
    $Nuget = Install-Package Nuget –force
    $NuGetModuleVersion = $($Nuget).version
    Write-Host ("Nuget [{0}] is installed" -f $NuGetModuleVersion) -ForegroundColor Green
    
    #get path to nuget provider
    $PackageManagementSourcePath = Get-ChildItem "$env:ProgramFiles\PackageManagement" -Recurse -Filter *.dll
    #build destingation path for backup
    $PackageManagementDestPath = "$BinPath\nuget\$PackageManagementAssemblyVersion"
    #test to see if same version exist in copied location
    $PackageManagementCopiedPath = Get-ChildItem $PackageManagementDestPath -Filter *.dll -Recurse -ErrorAction SilentlyContinue
    
    #If dll does NOT exist or refesh is enforced, copy files
    If($Refresh -or !($PackageManagementCopiedPath)){
        New-Item $PackageManagementDestPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Write-Host ("Copying nuget Assembly [{0}] from {1}" -f $PackageManagementAssemblyVersion,$PackageManagementSourcePath.FullName) -ForegroundColor Gray
        Copy-ItemWithProgress $PackageManagementSourcePath.FullName $PackageManagementDestPath -Force
    }
    Else
    {
        Write-Host ("Nuget Provider [{0}] already found" -f $PackageManagementAssemblyVersion) -ForegroundColor Gray
    }


    
    #loop through each module from list
    Foreach ($Module in $OnlineModules){
        
        #get the module if found online
        $OnlineModuleFound = Find-Module $Module

        If($OnlineModuleFound)
        {
            #grab the name and version of the online version
            [string]$ModuleName = $OnlineModuleFound.Name
            [string]$ModuleVersion = $OnlineModuleFound.Version
            
            #we want to standardized folder names for easier management
            #$ModuleBuildFolder = ($ModuleName + "-" + $ModuleVersion)
            $ModuleBuildFolder = "\b$ModuleName\b[^\n]+\b$ModuleVersion\b"

            #find modules with no simliarity; We can't use this; these will be removed if switched
            $NotLikeModulesExists = Get-ChildItem $ModulesPath -Directory | Where-Object {$_.Name -notmatch $ModuleName}
            
            #is there a module that has a package subfolder (from online nupkg)? We can't use this; these also will be removed if switched
            $NupkgModuleExits = Split-Path (Get-ChildItem -Path $ModulesPath -Recurse | Where-Object { $_.PSIsContainer} | Where-Object {$_.Name -eq "package"}).FullName -Parent

            #find modules with for exact match; We will use this
            $ExactModulesExists = Get-ChildItem $ModulesPath -Directory | Where-Object {$_.Name -eq "$ModuleName-$ModuleVersion"}

            #validate the module structure
            If($ExactModulesExists){
                $ValidModule = Inventory-ModuleStructure -ModulePath $ExactModulesExists.FullName -Validate
                #if specified we can try to fix the module structure
                If(!$ValidModule -and $FixInvalidModule){
                    $FixedModule = Inventory-ModuleStructure -ModulePath $ExactModulesExists.FullName -Fix 
                }
            }
            Else
            {
                #find modules with simliarity but exclude nupkg; We can use this
                $SimilarModulesExists = Get-ChildItem $ModulesPath -Directory | Where-Object {$_.Name -match $ModuleBuildFolder -and $_.FullName -ne $NupkgModuleExits -and $_.Name -ne $ExactModulesExists}
            
                #we don't know that the module is correctly formatted; this must be checked
                $FoundPotentialModule = Inventory-ModuleStructure -ModulePath $SimilarModulesExists.FullName -Validate -Verbose -DestRootPath $ModulesPath

                #After inventoring the module is it a potential match? 
                #If so fix it up for import
                $InventoriedModule = $FoundPotentialModule | Where {$_.ModuleType -eq "Manifest"}
                If(($InventoriedModule.RootModule + "." + $InventoriedModule.RootModuleVersion) -match $ModuleBuildFolder){
                    If($FixInvalidModule){
                        $FoundModule = Inventory-ModuleStructure -ModulePath $SimilarModulesExists.FullName -DestRootPath $ModulesPath -Fix
                    }
                    Else{
                        $FoundModule = $InventoriedModule 
                    }
                }
            }

            #determine if there are simliar modules but may not be named correctly; We might use this if parsed
            $LikeModulesExists = Get-ChildItem $ModulesPath -Directory | Where-Object {$_.Name -match $ModuleName -and $_.Name -notmatch $ModuleBuildFolder}
            
            #double check like modules to ensure they are not just renamed wrong
            If(-not($SimilarModulesExists) -and $LikeModulesExists){
                foreach ($LikeModule in $LikeModulesExists) {
                    #we don't know that the module is correctly formated; this must be checke
                    $FoundModule = Inventory-ModuleStructure -ModulePath $SimilarModulesExists.FullName -DestRootPath $ModulesPath
                }

            }

            $SameModulesExists = Get-ChildItem $ModulesPath -Directory | Where-Object {$_.Name -match $ModuleBuildFolder -and $_.FullName -ne $NupkgModuleExits}

        }




        #If specified, remove older modules in DownloadedModule directory if found
        If($RemoveOld)
        {
            foreach ($Modules in $LikeModulesExists) {
                Remove-Item -Force -Recurse
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
        If($InstallAsWell)
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

}
Else{
    Write-Host "ERROR: Unable to connect to the internet to grab modules" -ForegroundColor Red
    throw $_.error
}
