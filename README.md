# PowerCLI Installer v2 (aka v3)
*** Installing PowerCLl Offline **

## FILE: Install-PowerCLI.ps1
<b>URL:</b> http://www.powershellcrack.com/2017/09/installing-powercli-on-disconnected.html<br />
<b>DESCRIPTION:</b>Install an Offline copy of PowerCLI from https://www.powershellgallery.com/packages/VMware.PowerCLI
loading all dependencies in the appropiate order. Also installs the nuget assemby

<b>FEATURES:</b>
<li>Delete any prior PowerCLI's installed</li>
<li>Find the latest downloaded PowerCLI version folder, and copy the files to the users profile Powershell directory
Include a progress bar for coping</li>
<li>Copy Nuget folder (located in PowerCLI version folder) to C:\Program Files\PackageManagement\ProviderAssemblies</li>
<li>Disable CEIP</li>
<li>If switched, copy a modified version of Initialize-PowerCLIEnvironment.ps1 to the users profile PowerCLI directory</li>
<li>If switched, Copy a icon to the users profile PowerCLI directory</li>
<li>If switched, Create PowerCLI desktop shortcuts, x86 and x64 (like the installer did, pointing to Initialize-PowerCLIEnvironment.ps1  and using the icon)</li>
<li>Load PowerCLI when done. </li>

## How to use:
 1. Clone or downlaod a copy of this project
 2. On an internet connected device:<br />
    run: 
             
        powershell.exe -ExecutionPolicy Bypass -file "Install-PowerCLI.ps1"
              
    MANUALLY: <br />      
     Go to https://code.vmware.com/web/dp/tool/vmware-powercli --> Download the "VMware-PowerCLI-<version>.zip" and extract it to this modules location<br />
     or Save-Module VMware.PowerCli -Path <location of this module folder><br />
 2. Copy entire PowerCLI-ModuleInstaller to disconnected system<br />
 3. Run: <br />
 
        powershell.exe -ExecutionPolicy Bypass -file "Install-PowerCLI.ps1" -CreateShortcut
