# PowerCLI Installer v2 (aka v3)
*** Installing PowerCLl Offline **

##FILE: Install-PowerCLI.ps1
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

<hr>
<b>FILE:</b> Initialize-PowerCLIEnvironment.ps1<br />
<b>DESCRIPTION:</b>Modified version of VMWare's ps1 file that comes with 6.5.0 installer to be used with 6.5.2. 
This script is called by the desktop shortcuts to load the PowerCLI modules appropiately

