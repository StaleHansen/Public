<#
---------------------------------------------------
Created by MVP Ståle Hansen
---------------------------------------------------

Get-InstalledLyncVersion.ps1 script to check patch level on all Lync servers in an environment

.Notes
   - This script works for Lync Server 2010 and Lync Server 2013
   - Run Lync Management Shell in Administrative mode
        - Use the Function Get-RemoteProgram by Jaap Brasser, modified by MVP Jan Egil Ring
        - If you can not connect to a server make sure Remote Registry Service is running on the machine you are unable to reach ◾Port 139 and 445 needs to be open on the server you are trying to scan
   - Tested OK in 2010 and 2013 environments
   - The script will be updated as new CU’s will arrive, but will also flag unknown CU versions and let you download the latest one

    V1.0 May 2014 – Initial Script ◦Scans registry on all servers in a Lync deployment to find current installed version
        - You shall not use WMI to query MSI packets as that will invoke a reconfiguraton 
            - Read more here: http://powershell.org/wp/forums/topic/alternatives-to-win32_product/
        - Checks against a database in this script to find the correct CU level
        - Will ask if you want to download the current CU if there are servers which is not on the latest patch level 
            - Using function New-FileDownload by MVP Pat Richard 
            - Will open the download location after the file is downloaded
        - No Parameters are required
    V1.1 August 2014 - Added Lync Server 2013 Cumulative Update (CU5)
    V1.2 September 2014 - Added Lync Server 2013 Cumulative Update (CU6)
    V1.3 November and December updates added

.Link
   Twitter: http://www.twitter.com/StaleHansen
   Blogpost: http://wp.me/pv8hB-19K
   LinkedIn: http://www.linkedin.com/in/StaleHansen
   Current Release: V1.0
.EXAMPLE
   Get-InstalledLyncVersion.ps1
   Description:
   Will run the script and all of its content with the functions 
#>

########################################################
##         Lync Cumulative Updates Database           ##
##      as started by Max Sanna (max@maxsanna.com)    ##
########################################################

# Add further lines to the array as new CUs come out
$CUs = @()
$CUs += ,@("Lync 2010","November 2010 RTM","4.0.7577.0","NotLatest")
$CUs += ,@("Lync 2010","January 2011 Update (CU1)","4.0.7577.108","NotUpToDate")
$CUs += ,@("Lync 2010","April 2011 Update (CU2)","4.0.7577.137","NotUpToDate")
$CUs += ,@("Lync 2010","July 2011 Update (CU3)","4.0.7577.166","NotUpToDate")
$CUs += ,@("Lync 2010","November 2011 Update (CU4)","4.0.7577.183","NotUpToDate")
$CUs += ,@("Lync 2010","February 2012 Update (CU5)","4.0.7577.190","NotUpToDate")
$CUs += ,@("Lync 2010","June 2012 Update (CU6)","4.0.7577.199","NotUpToDate")
$CUs += ,@("Lync 2010","October 2012 Update (CU7)","4.0.7577.203","NotUpToDate")
$CUs += ,@("Lync 2010","March 2013 Update (CU8)","4.0.7577.216","NotUpToDate")
$CUs += ,@("Lync 2010","July 2013 Update (CU9)","4.0.7577.217","NotUpToDate")
$CUs += ,@("Lync 2010","October 2013 Update (CU10)","4.0.7577.223","NotUpToDate")
$CUs += ,@("Lync 2010","January 2014 Update (CU11)","4.0.7577.225","NotUpToDate")
$CUs += ,@("Lync 2010","April 2014 Update (CU12)","4.0.7577.230","UpToDate")
$CUs += ,@("Lync 2013","November 2012 RTM","5.0.8308.0","NotUpToDate")
$CUs += ,@("Lync 2013","February 2013 (CU1)","5.0.8308.291","NotUpToDate")
$CUs += ,@("Lync 2013","July 2013 Update (CU2)","5.0.8308.420","NotUpToDate")
$CUs += ,@("Lync 2013","October 2013 Update (CU3)","5.0.8308.556","NotUpToDate")
$CUs += ,@("Lync 2013","January 2014 Update (CU4)","5.0.8308.577","NotUpToDate")
$CUs += ,@("Lync 2013","August 2014 Update (CU5)","5.0.8308.738","NotUpToDate")
$CUs += ,@("Lync 2013","September 2014 Update (CU6)","5.0.8308.815","NotUpToDate")
$CUs += ,@("Lync 2013","November 2014 Update (CU7)","5.0.8308.834","NotUpToDate")
$CUs += ,@("Lync 2013","December 2014 Update (CU8)","5.0.8308.857","UpToDate")

#Get-RemoteProgram Function by Jaap Brasser and modified by MVP Jan Egil Ring
Function Get-RemoteProgram {
    param(
        [CmdletBinding()]
        [string[]]$ComputerName = $env:COMPUTERNAME
    )
    foreach ($Computer in $ComputerName) {
        
        try {
        $RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
        $RegUninstall = $RegBase.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
        $RegUninstall.GetSubKeyNames() | 
        ForEach-Object {
            $DisplayName = ($RegBase.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$_")).GetValue('DisplayName')
            $DisplayVersion = ($RegBase.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$_")).GetValue('DisplayVersion')
            if ($DisplayName -match "Core components") {
                New-Object -TypeName PSCustomObject -Property @{
                    ComputerName = $Computer
                    PoolName = (Get-CsComputer $Computer | Select-Object -ExpandProperty Pool)
                    ProgramName = $DisplayName
                    InstalledVersion = $DisplayVersion
			        ServerRole = $Null
                    IsUpToDate = $Null
                    
                }
            
            }
            }
        }
        catch {
        #"Failed to connect to $computer"
                    New-Object -TypeName PSCustomObject -Property @{
                    PoolName = (Get-CsComputer $Computer | Select-Object -ExpandProperty Pool)
                    ComputerName = $Computer
                    Connection = "Unable to connect to computer"
                    ServerRole = $Null
           
                }
        }
    }
}
#Start of script

CLS
Write-Output "Scanning all Lync servers..."

$Output=Get-RemoteProgram -ComputerName (Get-CsPool | Select-Object -ExpandProperty Computers)
    $IsLync2013=$False
    $IsLync2010=$False
    $Latest=0
    #Checking version on all computers we where able to connect to
    foreach ($Line in $Output){
	    if ($Line.ProgramName -match "Core"){
            $Service=Get-CsPool (Get-CsComputer $Line.ComputerName | Select-Object -ExpandProperty Pool) | Select-Object -expandproperty services
	        $Service=$Service | ForEach-Object {$_.split(':')[0]}
            if ($Service -match "UserServer"){$Service="Front End"}
            elseif ($Service -match "EdgeServer"){$Service="Edge"}
            elseif ($Service -match "UserDatabase"){$Service="SQL Backend"}
            $Line.ServerRole=$Service
            foreach ($CU in $CUs){
		        if ($CU -match $Line.InstalledVersion){
			    $Line.InstalledVersion=$CU[2]+" - "+$CU[1]
                $Line.IsUpToDate=$CU[3]
			    $Found = "CorrectVersionFound"
                if ($CU[3] -eq "NotUpToDate"){$Latest++}
		        }
            }
            if (!($Found -match "CorrectVersionFound")){$Line.InstalledVersion=$Line.InstalledVersion+" - Unknown CU found";$Latest++}
            if ($Line.FreeDiskSpace -ne $Null){$Line | fl PoolName, ComputerName, ProgramName, InstalledVersion, IsUpToDate, ServerRole. FreeDiskpace}
            else {$Line | fl PoolName, ComputerName, ProgramName, InstalledVersion, IsUpToDate, ServerRole}
            if ($Line.InstalledVersion -Like "5.0*"){$IsLync2013=$True}
            if ($Line.InstalledVersion -Like "4.0*"){$IsLync2010=$True}
        }

    }
    #Going through all cumputers we where unable to connect to and also ignoring PSTN gateways
    foreach ($Line in $Output){
        if ($Line.Connection -match "Unable"){
            if ((Get-CsPool (Get-CsComputer $Line.ComputerName | Select-Object -ExpandProperty Pool) | Select-Object -expandproperty services) -match "pstn"){}
	     
		    else {$Service=Get-CsPool (Get-CsComputer $Line.ComputerName | Select-Object -ExpandProperty Pool) |  Select-Object -expandproperty services
	            $Service=$Service | ForEach-Object {$_.split(':')[0]}
	            if ($Service -match "UserServer"){$Service="Front End"}
            	    elseif ($Service -match "EdgeServer"){$Service="Edge"}
            	    elseif ($Service -match "UserDatabase"){$Service="SQL Backend"}
            	    $Line.ServerRole=$Service
            	    $Line | fl  Connection, PoolName, ComputerName, ServerRole
        	}
	    }
        elseif ($Line.Connection -eq "$Null"){
            $Service=Get-CsPool (Get-CsComputer $Line.ComputerName | Select-Object -ExpandProperty Pool) | Where-Object {!($_.Services -match "pstn")} | Select-Object -expandproperty services
	        $Service=$Service | ForEach-Object {$_.split(':')[0]}
	        if ($Service -match "UserServer"){$Service="Front End"}
            elseif ($Service -match "EdgeServer"){$Service="Edge"}
            elseif ($Service -match "UserDatabase"){$Service="SQL Backend"}
            $Line.ServerRole=$Service
            $Line | fl  Connection, PoolName, ComputerName, ServerRole
        }
    }

#New file downlaod function by MVP Pat Richard (@patrichard)
function New-FileDownload {
	#[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)] 
		[ValidateNotNullOrEmpty()]
		[string]$SourceFile,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)] 
		[string]$DestFolder,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)] 
		[string]$DestFile,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)] 
		[switch]$IgnoreLocalCopy=$True
	)
	[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)

	if (!($DestFolder)){
		$DestFolder = $TargetFolder
	}
	Import-Module -name BitsTransfer
	#if (!($DestFile)){
	#	[string] $DestFile = ($SourceFile | Split-Path -leaf)
	#}
	if (Test-Path $DestFolder){
		Write-Output "Target folder `"$DestFolder`" exists - no need to create"
	} else {
		Write-Output "Folder `"$DestFolder`" does not exist, creating..."
		New-Item $DestFolder -type Directory | Out-Null
		Write-Output "Done!" 
	}
	if ((Test-Path "$DestFolder\$DestFile") -and (! $IgnoreLocalCopy)){
		Write-Output "File `"$DestFile`" exists locally - no need to download"
	} else {
		if ($HasInternetAccess){
			Write-Output "Internet access available"
			if (! $IgnoreLocalCopy){
				Write-Output "File `"$DestFile`" does not exist in `"$DestFolder`""
			} else {
				Write-Output "Forcing download of `"$DestFile`" to `"$DestFolder`""
			}
			Write-Output "Downloading `"$SourceFile`" to `"$DestFolder`""
			########################################################################################################
			# NOTE: Default parameters may have been changed due to proxy settings. See Test-IsProxyEnabled function
			########################################################################################################
			Start-BitsTransfer -Source $SourceFile -Destination "$DestFolder\$DestFile" -ErrorAction SilentlyContinue
			if (Test-Path $DestFolder\$DestFile){
				Write-Output "Done!"
			} else {
				Write-Output -level error -message "Failed! File not downloaded!"
				Write-Output "Prompting user to exit"
				If ((Read-Host "A file download failure has occured. Would you like to exit the script") -imatch "y"){
					Write-Output "User has chosen to exit script"
					Write-Output "Exiting script"
					exit
				}
			}
		} else {
			Write-Output -level warn -Message "Internet access not detected. Please resolve and try again."
		}
	}
} # end function New-FileDownload

#Download the latest CU
if ($latest -gt 0){
    $opt=Read-Host There is a newer Cumulative Update available, press 1 to download latest version. Press 9 to exit
    if ($opt -eq 1){
        if($IsLync2013 -eq $True){
            Write-Output "Downloading Lync Server 2013 Latest Cumulative Update"
            Write-Warning "Remember, you need at least 20 GB of free diskspace on the STD edition server or Enterprise SQL Back End to be able to update Lync database schema"
	        New-FileDownload -SourceFile "http://download.microsoft.com/download/B/E/4/BE44AC91-C665-4522-BA93-CE72B0934DAF/LyncServerUpdateInstaller.exe" -DestFolder C:\_Install -DestFile LyncServerUpdateInstaller2013.exe
            Invoke-Item c:\_install\
	    }
    	if($IsLync2010 -eq $True){
        	Write-Output "Downloading Lync Server 2010 Latest Cumulative Update"
        	New-FileDownload -SourceFile "http://download.microsoft.com/download/3/4/1/341C256C-0E74-4968-B6FA-EEA87600E283/LyncServerUpdateInstaller.exe" -DestFolder C:\_Install -DestFile LyncServerUpdateInstaller2010.exe
	        Invoke-Item c:\_install\
        }
    }
    else{}
}