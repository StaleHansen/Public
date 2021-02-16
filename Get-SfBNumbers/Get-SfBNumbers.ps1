<#
.SYNOPSIS
This script will get the next available number of any provided number range from Unassigned Numbers, Array or input to the script and will generate a full report and a summary per range

.NOTES
V1.0 - Added and verified the $ReportPieChartHTML paramenter and set it to $True as default
     - Fixed bug on connecting to SQL monitoring CDR databases with named instances and mirrored SQL backends
     - Optimized adding user activity to internal database and added parameter $ReportUserActivity that is default set to $False which will skip it since it may take up to 20 minutes
     - Added and verified the option to autmoatically classify numbers to Gold and Silver with parameter $AutoclassifyNumbers which default to $True
     - added option to classify Bronze number with parameter $ReserveBronzeNumbers which defaults to $False, do not run
     - Added check if Server 2010 to skip Get-CsMeetingRooms as the cmdlet does not exist
     - Added option to specify extension length with parameter $ExtensionLength that defaults to 43
V0.9 - Initial script as demoed at Microsoft Ignite 2015, the script is still en beta stage, but tested in multiple deployments

.DESCRIPTION
This script will get the next available number of any provided number range from
    •Unassinged Numbers
    •array in the script
    •parameter input when running the script
It will check for
    •disabled users in Active Directory
    •connect to the LcsCDR to get users that has not logged on for 30 days or more
    •Connect to LcsCDR to check for acitivty on numbers, both unassigned and assigned
It supports
    •Lync Server 2010, Lync Server 2013, Skype for Business Server 2015
    •Reservation of numbers based on categories like Gold, Silver, Bronze, Special and Excluded
        ◦From Array in the script or Unassigned Numbers
    •Reservations of number ranges within larger number ranges for special services like future callcenter or response groups
    •Retention of numbers based on disabled in Activer Directory, last logon time and activity on numbers
Reporting supported
    •Default behaviour is GridView with all numbers and summary in the PowerShell window
    •Export to of all information to CSV
    •Create a summary in PieChart as HTML


.PARAMETER
Parameters will get documented in detail at a later time

.EXAMPLE
.\Get-SfBNumbers.ps1
No parameters will do all series, unassigned numbers and output to gridview

.EXAMPLE
.\Get-SfBNumbers.ps1 -ReportGrid $True -Name MainRange
get all info on specific range and output to GridView

.EXAMPLE
.\Get-SfBNumbers.ps1 -AllRanges $False -ReportGrid $False
Choose range from a menu

.EXAMPLE
.\Get-SfBNumbers.ps1 -IncludeUnassignedNumbers $False
Output all ranges to gridview, ignotring Unassigned Number

.EXAMPLE
.\Get-SfBNumbers.ps1 -AllRanges $True -ReportGrid $False -ReportCSV $True
Output all ranges to csv c:\_reports\

.EXAMPLE
.\Get-SfBNumbers.ps1 -AllRanges $True -ReportGrid $True -ReportPieChartHTML $True
Output all ranges to piechart html stored in c:\_reports\

.EXAMPLE
.\Get-SfBNumbers.ps1 -AllRanges $false -ReportGrid $False -Name ExampleMainRange -FirstAvailableAsLineURI $True
Output all ranges to piechart html stored in c:\_reports\

.EXAMPLE
.\Get-SfBNumbers.ps1 -RangeStart +4721081655 -RangeEnd +4721081680 -Name Norway -AllRanges $false -ReportGrid $False
Use custom range and report only on that range

.EXAMPLE
#Return correct formatted LineUri from specified Range including extension
.\Get-SfBNumbers.ps1 -AllRanges $false -ReportGrid $False -Name ExampleMainRange -FirstAvailableAsLineURI $True

.INPUTS
The script does not support piped input at this time

.OUTPUTS
The script produces fully functional powershell output either you want the summary or use the $ReturnAllInfo $True switch to return not just a summary but all info as seen in the GridView

.LINK
http://SfBNumbers.net 

#>
[CmdletBinding()]
Param(
[Parameter(Mandatory=$False, ValueFromPipelineByPropertyName=$true)]
   $IncludeUnassignedNumbers=$True,
   $RangeStart=$Null,
   $RangeEnd,
   $Name="CustomRange",
   $AllRanges=$True,
   $FirstAvailable=$False,
   $FirstAvailableAsLineURI=$False,
   $ExtensionLength=4,
   $AvailableReserved,
   $ReturnAllInfo=$False,
   $ReportCSV=$False,
   $ReportHTML=$False,
   $ReportEmail=$False,
   $ReportGrid=$True,
   $ReportUserActivity=$False,
   $ReportPieChartHTML=$True,
   $AutoclassifyNumbers=$True,
   $ReserveBronzeNumbers=$False

)

#Global Variables
$AvailableNumbers=$Null

########################################################
##             Phone Number Range Database
##          Add your custom ranges in advance
########################################################
#Here you can add all your custom number ranges
$CRs = @()
$CRs += ,@("NMBU","+4767230000","+4767232999")

########################################################
##           Phone Number Retention Database
##      For your Gold, silver and special numbers
########################################################
#Here you can add all your custom number ranges
$NRs = @()
$NRs += ,@("Gold","+442030021880")

###############################################################################
## STORING ALL AVAILABLE RANGES IN AN OBJECT
###############################################################################

#create Object for all Number series
Function Add-AllSeries {
    New-Object -TypeName PSCustomObject -Property @{
        Identity = $Null
        NumberRangeStart = $Null
        NumberRangeEnd = $Null
    }
}

#adding all unassigned number series
$AllSeries = @()
$ReservedSeries = @()
$addObject = Add-AllSeries

if ($IncludeUnassignedNumbers -eq $True){
    foreach ($Range in (Get-CsUnassignedNumber)){
        if (($Range.NumberRangeEnd.Substring(5) - $Range.NumberRangeStart.Substring(5)) -gt 5000){
            $addObject = Add-AllSeries
            $addObject.Identity = $Range.Identity+" Range is over 5000 and is excluded"
            $addObject.NumberRangeStart = $Null
            $addObject.NumberRangeEnd = $Null
            $AllSeries += $addObject
        }

        if ($Range.Identity -match "Res"){
            $addObject = Add-AllSeries
            $addObject.Identity = $Range.Identity.Substring(4)
            $addObject.NumberRangeStart = $Range.NumberRangeStart.Substring(4)
            $addObject.NumberRangeEnd = $Range.NumberRangeEnd.Substring(4)
            $ReservedSeries += $addObject
            continue
        } 
        if ($Range.Identity -match "Diamond"){continue}
        if ($Range.Identity -match "Gold"){continue} 
        if ($Range.Identity -match "Silver"){continue} 
        if ($Range.Identity -match "Bronze"){continue} 
        if ($Range.Identity -match "Special"){continue}
        if ($Range.Identity -match "Excluded"){continue}

        
        
        else {
            $addObject = Add-AllSeries
            $addObject.Identity = $Range.Identity
            $addObject.NumberRangeStart = $Range.NumberRangeStart.Substring(4)
            $addObject.NumberRangeEnd = $Range.NumberRangeEnd.Substring(4)
            $AllSeries += $addObject
        }
    }
}

#Adding predefined custom ranges
if ($CRs -ne $Null){
    foreach ($Range in $CRs){
        
        if (($Range[1].Substring(1) - $Range[2].Substring(1)) -gt 5000){
            $addObject = Add-AllSeries
            $addObject.Identity = $Range[0]+" Range is over 5000 and is excluded"
            $addObject.NumberRangeStart = $Null
            $addObject.NumberRangeEnd = $Null
            $AllSeries += $addObject
        }
        
        if ($Range[0] -match "Res"){
            $addObject = Add-AllSeries
            $addObject.Identity = $Range[0].Substring(4)
            $addObject.NumberRangeStart = $Range[1]
            $addObject.NumberRangeEnd = $Range[2]
            $ReservedSeries += $addObject
        } 
        
        else {
            $addObject = Add-AllSeries
            $addObject.Identity = $Range[0]
            $addObject.NumberRangeStart = $Range[1]
            $addObject.NumberRangeEnd = $Range[2]
            $AllSeries += $addObject
        }

    }
}

#Adding the provided custom range
if ($RangeStart -ne $Null){
        
        if (($RangeStart.Substring(1) - $RangeEnd.Substring(1)) -gt 5000){
            $addObject = Add-AllSeries
            $addObject.Identity = $Name+" Range is over 5000 and is excluded"
            $addObject.NumberRangeStart = $Null
            $addObject.NumberRangeEnd = $Null
            $AllSeries += $addObject
        }
        
        else {
            $addObject = Add-AllSeries
            $addObject.Identity = $Name
            $addObject.NumberRangeStart = $RangeStart
            $addObject.NumberRangeEnd = $RangeEnd
            $AllSeries += $addObject
        }


}


###############################################################################
## GET ALL THE DIDS ASSIGNED ANYWHERE IN LYNC
## Based on @paulvaillant's Get-LyncNumbers.ps1
## Check out his blogpost: http://paul.vaillant.ca/2015/03/18/managing-your-lync-phone-numbers.html  
###############################################################################

function NewLyncNumber {
	param($Type,$LineUri,$Name,$SipAddress,$Identity,$VoicePolicy)

	# clean up LineUri and look it up in $availableDids to see if it's "on trunk"
	$OnTrunk = $false
	$did = $null
    $ext = $null
	# parse the uri; drop the schema (tel:) and take only the main part split on ext
	if($LineUri -match '^tel:\+') {
		[long]$did = $LineUri.Substring(5) -split ';' | select -first 1
		$NumberRange = $availableDids -contains $did
	}
	if($LineUri -match 'ext=') {
		[long]$ext = $LineUri.Substring(5) -split '=' | select -Last 1
	}
	[pscustomobject]@{Type = $Type; LineURI = $LineUri; DisplayName = $Name; SipAddress = $SipAddress; Identity = $Identity; OnTrunk = $OnTrunk; DID = $did; ext = $ext;  NumberRange = $null; VoicePolicy = $VoicePolicy}
}
function NewLyncNumberFromAdContact {
	param($Type,$Contact)
	NewLyncNumber $Type $Contact.LineURI $Contact.DisplayName $Contact.SipAddress $Contact.Identity $Contact.VoicePolicy
}

Function Get-AllnumbersInDeployment{
    # Microsoft.Rtc.Management.ADConnect.Schema.ADUser
    $userUris = Get-CsUser -Filter {LineURI -ne $Null} -WarningAction SilentlyContinue | % { NewLyncNumberFromAdContact "User" $_ }
    $plUris = Get-CsUser -Filter {PrivateLine -ne $Null} -WarningAction SilentlyContinue | % { NewLyncNumberFromAdContact "User-PrivateLine" $_ }

    # Microsoft.Rtc.Management.ADConnect.Schema.OCSADAnalogDeviceContact
    $analogUris = Get-CsAnalogDevice -Filter {LineURI -ne $Null} -WarningAction SilentlyContinue | % { NewLyncNumberFromAdContact "AnalogDevice" $_ }

    # Microsoft.Rtc.Management.ADConnect.Schema.OCSADCommonAreaPhoneContact
    $caUris = Get-CsCommonAreaPhone -Filter {LineURI -ne $Null} -WarningAction SilentlyContinue | % { NewLyncNumberFromAdContact "CommonAreaPhone" $_ }

    # Microsoft.Rtc.Rgs.Management.WritableSettings.Workflow
    $rgsUris = Get-CsRgsWorkflow | ?{ $_.lineuri } -WarningAction SilentlyContinue | % { 
	    NewLyncNumber "RgsWorkflow" $_.LineURI $_.Name $_.PrimaryUri $_.Identity
    }

    # Microsoft.Rtc.Management.Xds.AccessNumber
    $dialinUris = Get-CsDialInConferencingAccessNumber -Filter {LineURI -ne $Null} -WarningAction SilentlyContinue | % { 
	    NewLyncNumber "DialInConferencingAccessNumber" $_.LineURI $_.DisplayName $_.PrimaryUri $_.Identity
    }

    # Microsoft.Rtc.Management.ADConnect.Schema.OCSADExUmContact
    $exumUris = Get-CsExUmContact -Filter {LineURI -ne $Null} -WarningAction SilentlyContinue | % { NewLyncNumberFromAdContact "ExUmContact" $_ }

    # Microsoft.Rtc.Management.ADConnect.Schema.OCSADApplicationContact
    $tepUris = Get-CsTrustedApplicationEndpoint -Filter {LineURI -ne $Null} -WarningAction SilentlyContinue | % { NewLyncNumberFromAdContact "TrustedApplicationEndpoint" $_ }

    # Microsoft.Rtc.Management.ADConnect.Schema.OCSADMeetingRoom
    # Skip if Lync Server 2010
    if ((Get-CsServerVersion) -notmatch "Server 2010"){
        #write-Host "Adding MeetingRooms"
        $lrsUris = Get-CsMeetingRoom -Filter {LineURI -ne $Null} -WarningAction SilentlyContinue | % { NewLyncNumberFromAdContact "MeetingRoom" $_ }
    }
    # combine all results together
    $allUsedNumbers = New-Object System.Collections.ArrayList 
    foreach($list in @($userUris,$plUris,$analogUris,$caUris,$rgsUris,$dialinUris,$exumUris,$tepUris,$lrsUris)) {
	    if($list -and $list.Length -gt 0) {
		    $allUsedNumbers.AddRange($list)
	    }
    }

    Return $allUsedNumbers
}

###############################################################################
## Finding the range we want to work with
###############################################################################

function Write-Menu{
    CLS
    Write-Warning 'For a more complete report use .\Get-CsPhoneNumbers.ps1 -AllRanges $True -IncludeUnassignedNumbers $True -ReportGrid $True'
    Write-Host
    Write-Host "Choose which number series you want to interact with"

    
    $Tempseries=@()
    $Tempseries += $Allseries
    $Tempseries += $ReservedSeries
    $menu = @{}
    for ($i=1;$i -le  $Tempseries.count; $i++) {
        Write-Host "$i. $( $Tempseries[$i-1].Identity)"
        $menu.Add($i,( $Tempseries[$i-1].Identity))
        }
    Write-Host 
    Write-Host "$i. Quit"
    Write-Host 
    $menu.Add($i,("Exit"))

    [int]$ans = Read-Host 'Enter selection'
    $selection = $menu.Item($ans)
    if ($ans -ge $i){exit}
    $SelectedRange = $Tempseries | where-object {$_.Identity -match $selection}
    Return $SelectedRange
}



Function Get-ActiveRange{
$Selected= Write-Menu
$Selected
$ActiveRange = $AllSeries | Where-Object {$_.identity -match "$Selected"}
$ActiveRange
}


#Find out what number range to work with
function Get-WorkingSeries{
    $Tempseries=@()
    $Tempseries += $Allseries
    $Tempseries += $ReservedSeries

    if ($AllRanges -eq $False -and $Name -ne "CustomRange"){
        $Range = $Tempseries | Where-Object {$_.identity -eq "$Name"}
    }


    if ($AllRanges -eq $False -and $RangeStart -eq $Null -and $Name -eq "CustomRange"){
    
        $Range = Get-ActiveRange
    }
    
    if ($AllRanges -eq $True){
        $Range = $Allseries
    }

return $Range

}

###############################################################################
## Collecting all available numbers based on working ranges
###############################################################################

function Add-AllnumbersinRange{

    New-Object -TypeName PSCustomObject -Property @{
        Identity = $null
        NumberInRange = $null

    }

}

function Get-AllAvailableNumbers{
    param($FunctionRange)

    $AllNumbersInRange= @()
    
    foreach ($Ser in $FunctionRange){
        if ($Ser.NumberRangeStart -ne $Null){
            
            [Long]$Start = $Ser.NumberRangeStart.Substring(1)
            [Long]$End = $Ser.NumberRangeEnd.Substring(1)

            while ($Start -lt ($End+1)){
                $addObject = Add-AllnumbersinRange
                $addObject.Identity = $Ser.Identity
                $addObject.NumberInRange = $Start
                $AllNumbersInRange += $addObject
                $Start++
            }
        }
    }

    Return $AllNumbersInRange
}

###############################################################################
## Create an object that has all users, used and available numbers
###############################################################################

function Add-CompleteWorkingRange{

    New-Object -TypeName PSCustomObject -Property @{
        Type = $null
        LineURI = $null
        DisplayName = $null
        SipAddress = $null
        Identity = $null
        OnTrunk = $null
        DID = $null
        ext = $null
        VoicePolicy = $Null
        NumberRange = $null
        InRetention = $null
        RetentionType = $null
        OfflineMoreThan30Days = $null
        DuplicateExtentionFound = $False
        NumberActivity = $Null
        Comment = $null
    }

}

function Get-CompleteWorkingRange($FunctionAllUSed, $FunctionAvailable){

    #Find all users with a number within a number range
    $MeasureObject = @()
    foreach ($number in $FunctionAvailable){
        $Count = 0
        $FoundMatch = $False
        $CountEnd = $FunctionAllUSed.count
        while ($Count -lt ($CountEnd+1)){
            if ($number.NumberInRange -eq $FunctionAllUSed[$Count].did){
                $addObject = Add-CompleteWorkingRange
                $addObject.Type = $FunctionAllUSed[$Count].Type
                $addObject.LineURI = $FunctionAllUSed[$Count].LineURI
                $addObject.DisplayName = $FunctionAllUSed[$Count].DisplayName
                $addObject.SipAddress = $FunctionAllUSed[$Count].SipAddress
                $addObject.OnTrunk = $FunctionAllUSed[$Count].OnTrunk
                $addObject.DID = '+'+$FunctionAllUSed[$Count].DID
                $addObject.Type = $FunctionAllUSed[$Count].Type
                $addObject.ext = $FunctionAllUSed[$Count].ext
                $addObject.VoicePolicy = $FunctionAllUSed[$Count].VoicePolicy
                $addObject.NumberRange = $number.Identity
                $addObject.InRetention = $False
                $addObject.RetentionType = $Null
                $addObject.OfflineMoreThan30Days = $Null
                $addObject.DuplicateExtentionFound = $False
                $addObject.NumberActivity = $Null
                $addObject.Comment = ''
                $MeasureObject += $addObject
                $count = $CountEnd
                $FoundMatch = $True
            }
            $Count++
        }
        #Adding the number that is not in use
        if ($FoundMatch -eq $False){
            $addObject = Add-CompleteWorkingRange
            $addObject.DID = '+'+$number.NumberInRange
            $addObject.NumberRange = $number.Identity
            $addObject.InRetention = $False
            $addObject.Comment = ''
            $MeasureObject += $addObject
        }
    }
        #Find all users with numbers not present in a number range
        $CountEnd = $MeasureObject.count
        foreach ($number in $FunctionAllUSed){
            $Count = 0
            $FoundMatch = $False
            while ($Count -lt ($CountEnd)){
                if ($MeasureObject[$Count].did -match $number.did){
      
                    $count = $CountEnd
                    $FoundMatch = $True

                }
                $Count++
            }

            if ($FoundMatch -eq $False){

                $addObject = Add-CompleteWorkingRange
                $addObject.Type = $number.Type
                $addObject.LineURI = $number.LineURI
                $addObject.DisplayName = $number.DisplayName
                $addObject.SipAddress = $number.SipAddress
                $addObject.OnTrunk = $number.OnTrunk
                $addObject.DID = '+'+$number.DID
                $addObject.Type = $number.Type
                $addObject.ext = $number.ext
                $addObject.InRetention = $False
                $addObject.RetentionType = $Null
                $addObject.OfflineMoreThan30Days = $Null
                $addObject.DuplicateExtentionFound = $False
                $addObject.NumberActivity = $Null
                $addObject.Comment = ''
                $MeasureObject += $addObject
            }
        }#>
    
    
    Return $MeasureObject
}


###############################################################################
## Checking for users that has not logged on for a long time
## 
## This function is based on Get-LyncOrphanedUsers-v0.2.1 
## by @GuyBachar, http://guybachar.us and @y0avb, http://y0av.me
###############################################################################

function add-SQLConnect{

    New-Object -TypeName PSCustomObject -Property @{
        ServerName = $null
        InstanceName = $null
    }

}

function Check-UserActivity($FunctionDays){

$SQLReportingServerInstances= @()
$ReturnValue = @()
$overallrecords = $null

#define number of days of inactivity
$DayInterval= $FunctionDays

$MonServerFound = Get-CsService | Where-Object {$_.Role -eq "MonitoringDatabase"}
            
$MonServer=$MonServerFound[0]
    
    $SQLConnect = [System.Data.Sql.SqlDataSourceEnumerator]::Instance.GetDataSources() | ? { $_.servername -eq $Monserver.PoolFqdn.Split(".")[0]} -ErrorAction Stop
    if ($SQLConnect -eq $Null){$InstanceParameter = "NonRequired"}

    $addObject = add-SQLConnect
    $addObject.ServerName = $MonServer.PoolFqdn.Split(".")[0]
    $addObject.InstanceName = $MonServer.SqlInstanceName
    $SQLReportingServerInstances = $addObject

    #region Create empty variable that will contain the user registration records
    $overallrecords = $null

    $SqlQuery = "SELECT Users.UserId,UserUri,LastLogInTime,LastConfOrganizedTime,LastCallFailureTime,LastConfOrganizerCallFailureTime `
                            FROM [LcsCDR].[dbo].[UserStatistics] `
                            INNER JOIN [LcsCDR].[dbo].[Users] ON [LcsCDR].[dbo].[Users].UserId = [LcsCDR].[dbo].[UserStatistics].UserId `
                            WHERE LastLogInTime IS NOT NULL `
                            ORDER BY [LcsCDR].[dbo].[UserStatistics].LastLogInTime desc"

    #Defnie SQL Connection
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SQLServer = $Monserver.PoolFqdn.Split(".")[0]
    $SQLInstance = $MonServer.SqlInstanceName

    if ($InstanceParameter -eq "NonRequired")
            {
                $SqlConnection.ConnectionString = "Server = $SQLServer; Database = lcscdr; Integrated Security = True"
		    }else
            {
                $SqlConnection.ConnectionString = "Server = $SQLServer\$SQLInstance; Database = lcscdr; Integrated Security = True"
		    }
            #Write-Host $SqlConnection.ConnectionString 

    try   {
            #Define SQL Command     
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandText = $SqlQuery
            $SqlCmd.Connection = $SqlConnection

            #Get the results
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $SqlAdapter.SelectCommand = $SqlCmd
        
            $DataSet = New-Object System.Data.DataSet
            $SqlAdapter.Fill($DataSet)
        
            $SqlConnection.Close()
                  }
    catch {
            write-host "Error Conencting to local SQL service on $SQLReportingServer with instance $SQLReportingServerInstances, Please verify connectivity and permissions" -ForegroundColor Red
            Break
          }
 
    #Append the results to the reuslts from the previous servers
    $overallrecords = $DataSet.Tables[0]

    $DateToCompare = (Get-date).AddDays(-$DayInterval)
    $overallrecords = $overallrecords | Where-Object {$_.LastLogInTime -lt $DateToCompare} -ErrorAction Ignore
    #endregion

    #region Script Output Display
    $filedate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
    $ServicesFileName = $env:TEMP+"\LastLogonExport-"+$filedate+".csv"
    $ListUsers = @()
    $overallrecords | ForEach-Object{ 

        # save a reference to the current user
        $user = $_           

        $tspan=New-TimeSpan $user.LastLogInTime (Get-Date);
        $diffDays=($tspan).days;

        # comment out to add just the LastRegisterTime property
        $user | Add-Member -MemberType NoteProperty -Name "DaysSinceLastLogin" -Value ($diffDays)
        $ListUsers = $ListUsers + $user
    } 

    $ReturnValue +=$overallrecords | Select-Object UserUri,LastLogInTime,DaysSinceLastLogin,LastConfOrganizedTime,LastCallFailureTime,LastConfOrganizerCallFailureTime 




    foreach ($User in $ReturnValue ){

        $Count=0

        while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
            if ($CompleteWorkingRange[$Count].sipaddress -match $User.UserUri){
                
                $CompleteWorkingRange[$Count].OfflineMoreThan30Days = $User.DaysSinceLastLogin
                $count = $CompleteWorkingRange.GetUpperBound(0)

            }
            $Count++
            
        }
    }
Return $ReturnValue

}


###############################################################################
## Checking for activity on numbers 
## 
## This function is based on Get-LyncOrphanedUsers-v0.2.1 
## by @GuyBachar, http://guybachar.us and @y0avb, http://y0av.me
###############################################################################

function Check-NumberActivity($FunctionDays){

$SQLReportingServerInstances= @()
$ReturnValue = @()
$overallrecords = $null

#define number of days of inactivity
$DayInterval= $FunctionDays

$MonServerFound = Get-CsService | Where-Object {$_.Role -eq "MonitoringDatabase"}
            
$MonServer = $MonServerFound[0]
    
    $SQLConnect = [System.Data.Sql.SqlDataSourceEnumerator]::Instance.GetDataSources() | ? { $_.servername -eq $Monserver.PoolFqdn.Split(".")[0]} -ErrorAction Stop
    if ($SQLConnect -eq $Null){$InstanceParameter = "NonRequired"}

    $addObject = add-SQLConnect
    $addObject.ServerName = $MonServer.PoolFqdn.Split(".")[0]
    $addObject.InstanceName = $MonServer.SqlInstanceName
    $SQLReportingServerInstances = $addObject

    #region Create empty variable that will contain the user registration records
    $overallrecords = $null

    $SqlQuery = "SELECT SessionIdTime,PhoneUri `
                            FROM [LcsCDR].[dbo].[VoipDetails] `
                            INNER JOIN [LcsCDR].[dbo].[Phones] ON [LcsCDR].[dbo].[Phones].PhoneId = [LcsCDR].[dbo].[VoipDetails].ConnectedNumberId `
                            WHERE ConnectedNumberId IS NOT NULL AND ToGatewayId IS NULL "
                         

    #Defnie SQL Connection
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SQLServer = $Monserver.PoolFqdn.Split(".")[0]
    $SQLInstance = $MonServer.SqlInstanceName

    if ($InstanceParameter -eq "NonRequired")
            {
                $SqlConnection.ConnectionString = "Server = $SQLServer; Database = lcscdr; Integrated Security = True"
		    }else
            {
                $SqlConnection.ConnectionString = "Server = $SQLServer\$SQLInstance; Database = lcscdr; Integrated Security = True"
		    }
            #Write-Host $SqlConnection.ConnectionString 

    try   {
            #Define SQL Command     
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandText = $SqlQuery
            $SqlCmd.Connection = $SqlConnection

            #Get the results
            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $SqlAdapter.SelectCommand = $SqlCmd
        
            $DataSet = New-Object System.Data.DataSet
            $SqlAdapter.Fill($DataSet)
        
            $SqlConnection.Close()
                  }
    catch {
            write-host "Error Conencting to SQL service on $SQLServer with instance $SQLInstance, Please verify connectivity and permissions" -ForegroundColor Red
            Break
          }
 
    #Append the results to the reuslts from the previous servers
    $overallrecords = $DataSet.Tables[0]

    $DateToCompare = (Get-date).AddDays(-$DayInterval)
    $overallrecords = $overallrecords | Where-Object {$_.SessionIdTime  -gt $DateToCompare} -ErrorAction Ignore

    $ReturnValue +=$overallrecords
   

#Return $ReturnValue
        Write-Host " Done!" -ForegroundColor Green
        #$ReturnValue = $ReturnValue | Where-Object {$_.PhoneURI -like $CompleteWorkingRange.did}

        Write-Host "Adding user call activity to internal database, this may take 5-20 minutes..." -ForegroundColor Yellow -NoNewline
        $SkipValue=@(0)
        foreach ($Number in $ReturnValue ){
            $Count2=0
            foreach ($Value in $SkipValue){
         
                if ($Value -eq $Number.PhoneURI){$count2 = $CompleteWorkingRange.GetUpperBound(0)}
                #Write-Host "wrting $Value $Number"
             }  
        
            if ($count2 -eq $CompleteWorkingRange.GetUpperBound(0)){Continue}
            $Activity=@($ReturnValue | Where-Object {$_.PhoneURI -like $Number.PhoneURI}).count
            #write-Host $Activity $Number.PhoneURI
            $SkipValue += $Number.PhoneURI
        
                

            while ($Count2 -lt ($CompleteWorkingRange.GetUpperBound(0))){
                if ($CompleteWorkingRange[$Count2].did -Like $Number.PhoneURI){
                    $CompleteWorkingRange[$Count2].NumberActivity = $Activity
                    if ($CompleteWorkingRange[$Count2].sipaddress -eq $Null){
                        $CompleteWorkingRange[$Count2].InRetention = $True
                        $CompleteWorkingRange[$Count2].Comment = "InRetention because of activity on the number"
                    }
                    $count2 = $CompleteWorkingRange.GetUpperBound(0)
                    Continue                
                }
                $Count2++
            }
            

        
        }

    Return $ReturnValue

}


###############################################################################
## Checking for disabled users that have a LineURI and marking them for retention
###############################################################################

function Get-Retention {

$DisabledADUsers = Get-CsAdUser | ?{$_.UserAccountControl -match "AccountDisabled" -and $_.Enabled -eq $True} | Get-CsUser | Where-Object {$_.lineuri -match "Tel:"} | Select-Object name, sipaddress, lineuri

    foreach ($User in $DisabledADUsers){

        $Count=0

        while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
            if ($CompleteWorkingRange[$Count].lineuri -Like $User.lineuri){
                
                $CompleteWorkingRange[$Count].InRetention = $True
                $CompleteWorkingRange[$Count].RetentionType = "DisabledUser"
                $count = $CompleteWorkingRange.GetUpperBound(0)
            }
            $Count++

        }
    }

}

###############################################################################
## Adding Gold, Silver and Special Numbers to retention defined at the top of the script
## Also adding custom ranges that is reserved within an other range at the top of the script
###############################################################################

function Get-GoldRetention {

    foreach ($Number in $NRs ){
        $Number[1]

        $Count=0
       
        while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
            if ($CompleteWorkingRange[$Count].did -contains $Number[1]){
                $CompleteWorkingRange[$Count].did
                $CompleteWorkingRange[$Count].InRetention = $True
                $CompleteWorkingRange[$Count].RetentionType = $Number[0]
                $count = $CompleteWorkingRange.GetUpperBound(0)
               
            }
            $Count++

        }
    }
    foreach ($Number in $CRs ){
        if ($Number[0] -match "Res_"){
                   
            [Long]$Start = $Number[1].Split("+")[1]
            [Long]$End = $Number[2].Split("+")[1]
            $Identity = $Number[0].Split("_")[1]
            

            $Count=0
            $Test = @()
            $addObject = Add-AllSeries
            $addObject.Identity = $Identity
            $addObject.NumberRangeStart = "+"+$Start
            $addObject.NumberRangeEnd = "+"+$End
            $FunctionAddedRange += $addObject
            while ($Start -lt ($End+1)){

                $Count=0
                while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
                    if ($CompleteWorkingRange[$Count].did -match $Start ){
                        $CompleteWorkingRange[$Count].did
                        $CompleteWorkingRange[$Count].NumberRange = $Identity
                        $count = $CompleteWorkingRange.GetUpperBound(0)
               
                    }
                    $Count++
                }
                $Start++
            }



        }
    }
    Return $NRs
}

###############################################################################
## Classifying numbers based on Gold, Silver and Bronze
## Function is created fully by Paul Valiant @paulvaillant
## Check out the original article: http://paul.vaillant.ca/2015/05/11/classifying-phone-numbers.html
###############################################################################
function Get-PhoneNumberClass{
    [CmdletBinding()]
    param(
	    [Parameter(Mandatory=$true, ParameterSetName="cli", Position=0)][long]$Number,
        [Parameter(ParameterSetName="cli")][switch]$Details,
        [Parameter(ParameterSetName="pipeline", ValueFromPipeline = $true)][long[]]$Pipeline,
        [Parameter(Mandatory=$true, ParameterSetName="test")][switch]$Test,
        [Parameter(ParameterSetName="test")][int][ValidateRange(1,1000000)]$RunSize = 100,
        [Parameter()][switch]$Slow
    )

    BEGIN {
        # the only thing we do in begin is setup the regex rules we'll use below
        $CLASSES = @{
            Gold = @{
                doubleTriple = "(\d)\1(\d)\2{2}$";
                doubleDouble0 = "(\d)\1(\d)\2{1}0$";
                triple0 = "(\d)\1{2}0$";
                same4 = "(\d)\1{3}$";
                sequential4 = "(?:0(?=1)|1(?=2)|2(?=3)|3(?=4)|4(?=5)|5(?=6)|6(?=7)|7(?=8)|8(?=9)|9(?=0)){3}\d$"
            };
            Silver = @{
                double0 = "(\d)\1{1}0$";
                bond = "007$";
                twoDigitPattern = "(\d{2})\1$"
            };
            Bronze = @{
                double = "(\d)\1$";
                endsIn0 = "0$"
            }
        }

        $GOLD_RE = @()
        $GOLD_REASONS = @()
        $CLASSES["Gold"].GetEnumerator() | foreach {
            $GOLD_RE += $_.Value
            $GOLD_REASONS += $_.Key
        }

        $SILVER_RE = @()
        $SILVER_REASONS = @()
        $CLASSES["Silver"].GetEnumerator() | foreach {
            $SILVER_RE += $_.Value
            $SILVER_REASONS += $_.Key
        }

        $BRONZE_RE = @()
        $BRONZE_REASONS = @()
        $CLASSES["Bronze"].GetEnumerator() | foreach {
            $BRONZE_RE += $_.Value
            $BRONZE_REASONS += $_.Key
        }

        # this is an example of how to combine all the regex together into one and
        # still make it readable using white space, commenting and capture group names
        # SEE https://msdn.microsoft.com/en-us/library/yd1hzczs.aspx#Whitespace
        # because it's all combined, if you want to add to this regex you'll need to
        # make sure you update the \# placeholders as necessary
        $CLASS_RE = "(?x)
        (?:
            (?<Gold_doubleTriple>(\d)\1(\d)\2{2})
            |
            (?<Gold_doubleDouble0>(\d)\3(\d)\4 0)
            |
            (?<Gold_triple0>(\d)\5{2}0)
            |
            (?<Gold_same4>(\d)\6{3})
            |
            (?<Gold_sequential4>(?:0(?=1)|1(?=2)|2(?=3)|3(?=4)|4(?=5)|5(?=6)|6(?=7)|7(?=8)|8(?=9)|9(?=0)){3}\d)
            |
            (?<Silver_double0>(\d)\7 0)
            |
            (?<Silver_bond>007)
            |
            (?<Silver_twoDigitPattern>(\d{2})\8)
            |
            (?<Bronze_double>(\d)\9)
            |
            (?<Bronze_endsIn0>0)
        )$"

        function ClassifySlow($number) {
            Write-Verbose "Classifying $number (slow)"
            $class = ""
            $reason = ""
            for($i = 0; $i -lt $GOLD_RE.Length; $i++) {
                if($number -match $GOLD_RE[$i]) {
                    $class = "Gold"
                    $reason = $GOLD_REASONS[$i]
                    break
                }
            }
            if(!$class) {
                for($i = 0; $i -lt $SILVER_RE.Length; $i++) {
                    if($number -match $SILVER_RE[$i]) {
                        $class = "Silver"
                        $reason = $SILVER_REASONS[$i]
                        break
                    }
                }
            }
            if(!$class) {
                for($i = 0; $i -lt $BRONZE_RE.Length; $i++) {
                    if($number -match $BRONZE_RE[$i]) {
                        $class = "Bronze"
                        $reason = $BRONZE_REASONS[$i]
                        break
                    }
                }
            }
            if(!$class) {
                $class = "Ordinary"
            }
            @{Number = $number; Class = $class; Reason = $reason}
        }

        function ClassifyFast($number) {
            Write-Verbose "Classifying $number (fast)"
            $class = "Ordinary"
            $reason = ""
            if($number -match $CLASS_RE) {
                $class,$reason = $($matches.Keys | ? { $_ -notmatch "^[0-9]+$" }) -split '_'
            }
            @{Number = $number; Class = $class; Reason = $reason}
        }

        # we could do this another way, by having a function called Classify
        # that checked the $Slow parameter each time, but in the case of
        # a large number of piped values that would be dont once for each
        # value instead of how it's done below which is once per script call
        if($Script:Slow) {
            new-alias -Force -Scope Script -Name Classify -Value ClassifySlow
        } else {
            new-alias -Force -Scope Script -Name Classify -Value ClassifyFast
        }
    }

    PROCESS {
        if($PSCmdlet.ParameterSetName -eq "test")
        {
            # generate test numbers
            $numbers = 0..$RunSize | %{ Get-Random -Minimum 1991000000 -Maximum 9999999999 }

            Measure-Command { foreach($n in $numbers) { ClassifyFast $n | out-null } } |
                select @{n='Name';e={'Fast'}},TotalMilliseconds

            Measure-Command { foreach($n in $numbers) { ClassifySlow $n | out-null } } |
                select @{n='Name';e={'Slow'}},TotalMilliseconds
        }
        elseif($PSCmdlet.ParameterSetName -eq "pipeline")
        {
            Write-Verbose "Classifying values from pipeline"
            foreach($n in $Pipeline) {
                Classify $n
            }
        }
        else
        {
            Write-Verbose "Classifying value from cli"
            $c = Classify $Number
            if($Details) {
                $c
            } else {
                $c.Class
            }
        }
    }
}


###############################################################################
## Loop through all numbers and classify them as gold, silver and bronze
###############################################################################
function Start-NumberClassification {

foreach ($DID in $CompleteWorkingRange){

        $Class=Get-PhoneNumberClass $DID.did -Details
        $Class.number="+"+$Class.number
        $Class
        if($Class.Class -eq "Gold" -or $Class.Class -eq "Silver"){
            $Count=0
       
            while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
                if ($CompleteWorkingRange[$Count].did -contains $Class.number){
                    $CompleteWorkingRange[$Count].did
                    if ($CompleteWorkingRange[$Count].SipAddress -eq $Null){$CompleteWorkingRange[$Count].InRetention = $True}
                    $CompleteWorkingRange[$Count].RetentionType = $Class.Class
                    $count = $CompleteWorkingRange.GetUpperBound(0)
               
                }
                $Count++

            }
        }
    }

}

###############################################################################
## Checking to see if unassinged numbers are used as number retention
###############################################################################

function Get-UnassignedRetention {

    $Check = Get-CsUnassignedNumber
    $FunctionAddedRange=$AllSeries

    foreach ($RetentionUnassigned in $Check){

        if ($RetentionUnassigned.Identity -match "Diamond"){
            $Count=0
            while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
                if ($CompleteWorkingRange[$Count].did -contains (($RetentionUnassigned.NumberRangeStart).Split(":")[1])){
                    $CompleteWorkingRange[$Count].did
                    $CompleteWorkingRange[$Count].InRetention = $True
                    $CompleteWorkingRange[$Count].RetentionType = "Diamond"
                    $count = $CompleteWorkingRange.GetUpperBound(0)
               
                }
                $Count++
            }
        }
        if ($RetentionUnassigned.Identity -match "Gold"){
            $Count=0
            while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
                if ($CompleteWorkingRange[$Count].did -contains (($RetentionUnassigned.NumberRangeStart).Split(":")[1])){
                    $CompleteWorkingRange[$Count].did
                    $CompleteWorkingRange[$Count].InRetention = $True
                    $CompleteWorkingRange[$Count].RetentionType = "Gold"
                    $count = $CompleteWorkingRange.GetUpperBound(0)
               
                }
                $Count++
            }
        }
        if ($RetentionUnassigned.Identity -match "Silver"){
            $Count=0
            while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
                if ($CompleteWorkingRange[$Count].did -contains (($RetentionUnassigned.NumberRangeStart).Split(":")[1])){
                    $CompleteWorkingRange[$Count].did
                    $CompleteWorkingRange[$Count].InRetention = $True
                    $CompleteWorkingRange[$Count].RetentionType = "Silver"
                    $count = $CompleteWorkingRange.GetUpperBound(0)
               
                }
                $Count++
            }
        }
        if ($RetentionUnassigned.Identity -match "Bronze"){
            $Count=0
            while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
                if ($CompleteWorkingRange[$Count].did -contains (($RetentionUnassigned.NumberRangeStart).Split(":")[1])){
                    $CompleteWorkingRange[$Count].did
                    $CompleteWorkingRange[$Count].InRetention = $True
                    $CompleteWorkingRange[$Count].RetentionType = "Bronze"
                    $count = $CompleteWorkingRange.GetUpperBound(0)
               
                }
                $Count++
            }
        }
        if ($RetentionUnassigned.Identity -match "Special"){
            $Count=0
            while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
                if ($CompleteWorkingRange[$Count].did -contains (($RetentionUnassigned.NumberRangeStart).Split(":")[1])){
                    $CompleteWorkingRange[$Count].did
                    $CompleteWorkingRange[$Count].InRetention = $True
                    $CompleteWorkingRange[$Count].RetentionType = "Special"
                    $count = $CompleteWorkingRange.GetUpperBound(0)
               
                }
                $Count++
            }
        }
        if ($RetentionUnassigned.Identity -match "Excluded"){
            $Count=0
            while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
                if ($CompleteWorkingRange[$Count].did -contains (($RetentionUnassigned.NumberRangeStart).Split(":")[1])){
                    $CompleteWorkingRange[$Count].did
                    $CompleteWorkingRange[$Count].InRetention = $True
                    $CompleteWorkingRange[$Count].RetentionType = "Excluded"
                    $count = $CompleteWorkingRange.GetUpperBound(0)
               
                }
                $Count++
            }
        }
        if ($RetentionUnassigned.Identity -match "Res_"){
                   
            [Long]$Start = $RetentionUnassigned.NumberRangeStart.Split("+")[1]
            [Long]$End = $RetentionUnassigned.NumberRangeEnd.Split("+")[1]
            $Identity = $RetentionUnassigned.Identity.Split("_")[1]
            

            $Count=0
            $Test = @()
            $addObject = Add-AllSeries
            $addObject.Identity = $Identity
            $addObject.NumberRangeStart = "+"+$Start
            $addObject.NumberRangeEnd = "+"+$End
            $FunctionAddedRange += $addObject
            while ($Start -lt ($End+1)){

                $Count=0
                while ($Count -lt ($CompleteWorkingRange.GetUpperBound(0))){
                    if ($CompleteWorkingRange[$Count].did -match $Start ){
                        $CompleteWorkingRange[$Count].did
                        $CompleteWorkingRange[$Count].NumberRange = $Identity
                        $count = $CompleteWorkingRange.GetUpperBound(0)
               
                    }
                    $Count++
                }
                $Start++
            }

        }
 
        

    }


    Return $FunctionAddedRange
}

###############################################################################
## Reporting PieChats in HTML
###############################################################################

Function Report-PieChartHTML($FunctionReport){

New-Item -ErrorAction Ignore -ItemType directory -Path "c:\_Report\" 

Function Create-PieChart() {
       param($FunctionReport)

       $Filename="c:\_Report\"
              
       [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
       [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
       
       #Create our chart object 
       $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
       $Chart.Width = 600
       $Chart.Height = 400
       $Chart.Left = 200
       $Chart.Top = 200

       #Create a chartarea to draw on and add this to the chart 
       $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
       $Chart.ChartAreas.Add($ChartArea) 
       [void]$Chart.Series.Add("Data") 

       #Add a datapoint for each value specified in the arguments (args) 
  
              Write-Host "Now processing chart value: "$FunctionReport.Identity
              $datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $FunctionReport.NumberOfAvailableNumbers)
              $datapoint.AxisLabel = "NumberOfAvailableNumbers" + "(" + $FunctionReport.NumberOfAvailableNumbers + ")"
              $Chart.Series["Data"].Points.Add($datapoint)

              $UsedNumbers= ($FunctionReport.TotalNumbersInRange - $FunctionReport.NumberOfAvailableNumbers - $FunctionReport.TotalNumbersInRetention)
              if ($UsedNumbers -ne $Null){
              $datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $UsedNumbers)
              $datapoint.AxisLabel = "NumberOfUsedNumbers" + "(" + $UsedNumbers + ")"
              $Chart.Series["Data"].Points.Add($datapoint)
              }

              if ($FunctionReport.TotalNumbersInRetention -ne 0){
              $datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $FunctionReport.TotalNumbersInRetention)
              $datapoint.AxisLabel = "TotalNumbersInRetention" + "(" + $FunctionReport.TotalNumbersInRetention + ")"
              $Chart.Series["Data"].Points.Add($datapoint)
              }
    

       $Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
       $Chart.Series["Data"]["PieLabelStyle"] = "Outside" 
       $Chart.Series["Data"]["PieLineColor"] = "Black" 
       $Chart.Series["Data"]["PieDrawingStyle"] = "Concave" 
       ($Chart.Series["Data"].Points.FindMaxByValue())["Exploded"] = $true

       #Set the title of the Chart to the current date and time 
       $Title = new-object System.Windows.Forms.DataVisualization.Charting.Title 
       $Chart.Titles.Add($Title) 
       $Chart.Titles[0].Text = $FunctionReport.Identity +": Start " + $FunctionReport.NumberRangeStart + " - End " + $FunctionReport.NumberRangeEnd + " Total Numbers in range: " +$FunctionReport.TotalNumbersInRange

       #Save the chart to a file
       $Chart.SaveImage($FileName+$FunctionReport.Identity+".png","png")

       #$a = $a + $FunctionReport.Identity
       #$a = $a + "First number "+ $FunctionReport.NumberRangeStart + " last number " + $FunctionReport.NumberRangeEnd
       $Img= ($FileName+$FunctionReport.Identity+".png")
       $a = $a + @" 
       <html><body><br><img src="
"@
$a = $a + $Img 
$a = $a + @" 
"alt="
"@
$a = $a + $FunctionReport.Identity
$a = $a + @" 
" align="right"></body></html>
"@
Return $a

}

 $HTML = $Null
    foreach ($Rep in $FunctionReport){
        $HTML = $HTML + (Create-PieChart($Rep))
    }

    #$HTML
    Write-Host " Done!" -ForegroundColor Green
    $File = "C:\_Report\SfBNumbersPieCharts"+(get-date -f yyyy-MM-dd)+".htm"
    ConvertTo-HTML -PreContent $HTML | Out-File $File
    Invoke-Expression $File


}



###############################################################################
## Measuring and reporting available numbers
###############################################################################

Function Add-Report {
    New-Object -TypeName PSCustomObject -Property @{
        Identity = $Null
        NumberRangeStart = $Null
        NumberRangeEnd = $Null
        TotalNumbersInRange = $Null
        FirstavailableNumber = $Null
        FirstavailableExtension = $Null
        NumberOfAvailableNumbers = $Null
        TotalNumbersInRetention = $Null
        GoldnumbersInRetention = $Null
        SilverNumbersInRetention = $Null
        BronzeNumbersInRetention = $Null
        SpecialNumbersInRetention = $Null
        ExcludedNumbersInRetention = $Null
        DisabledUsersInRetention = $Null
        NumbersWithActivityInRetention = $Null
        Comment = $Null
    }
}

#Checking if extension is in use
function Get-Extension($FunctionRange){
    $Ext=$Null
    $ExtCheck=$Null
    $ArrayCount=0
    $Comment = ""
    $CountEnd = $CompleteWorkingRange.count
    $Count=$Null
    $Ext = ($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $FunctionRange -and $_.InRetention -eq $False})[$ArrayCount].did    
    $Ext = $Ext.Substring($Ext.Length-$ExtensionLength)
    $ExtCheck = $CompleteWorkingRange.ext
    if ($Ext -ne $Null){
        while ($Count -ne ($CountEnd)){
 

            foreach ($Check in $ExtCheck){
            
                if ($Check -eq $Ext){
                    $ArrayCount++
                       
                }
                $Count++
            }
        }
    }
    
    $Ext = ($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $FunctionRange -and $_.InRetention -eq $False})[$ArrayCount].did
    $Ext = $Ext.Substring($Ext.Length-$ExtensionLength)
    if ($ArrayCount -gt 0){$Comment="Previous Numbers were skipped because of conflicting extension"}
    if ($Ext -eq $Null){$ArrayCount -eq $Null}
    Return $Ext, $ArrayCount, $Comment

}

function Get-Report($FunctionCompleteRange){

    $ErrorActionPreference = 'SilentlyContinue'

    #Report statistics to Screen
    #if ($FirstAvailable -eq $False -and $ReportCSV -eq $False -and $ReportHTML -eq $False -and $ReportEmail -eq $False -and $ReportGrid -eq $False){
        $ReportToScreen = @()
        if($AllRanges -eq $True){$Tempseries += $AllSeries; $Tempseries += $ReservedSeries; $WorkingRange = $Tempseries}

        foreach ($Range in ($WorkingRange | Where-Object {$_.identity -ne $Null})){
                #write-host $Range.Identity
                $Extension=$Null
                $addObject = Add-Report
                $addObject.Identity = $Range.Identity
                $addObject.NumberRangeStart = $Range.NumberRangeStart
                $addObject.NumberRangeEnd = $Range.NumberRangeEnd
                $addObject.TotalNumbersInRange = @($CompleteWorkingRange | Where-Object {$_.NumberRange -eq $Range.Identity}).count               
                $addObject.NumberOfAvailableNumbers = @($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $Range.Identity -and $_.InRetention -eq $False} ).count           
                $Extension = Get-Extension $Range.Identity
                $addObject.FirstavailableExtension = $Extension[0]
                $addObject.FirstavailableNumber = ($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $Range.Identity -and $_.InRetention -eq $False})[$Extension[1]].did
                $addObject.GoldnumbersInRetention = @($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $Range.Identity -and $_.InRetention -eq $True -and $_.RetentionType -eq "Gold"}).count
                $addObject.SilverNumbersInRetention = @($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $Range.Identity -and $_.InRetention -eq $True -and $_.RetentionType -eq "Silver"}).count
                $addObject.BronzeNumbersInRetention = @($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $Range.Identity -and $_.InRetention -eq $True -and $_.RetentionType -eq "Bronze"}).count
                $addObject.SpecialNumbersInRetention = @($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $Range.Identity -and $_.InRetention -eq $True -and $_.RetentionType -eq "Special"}).count
                $addObject.DisabledUsersInRetention = @($CompleteWorkingRange | Where-Object {$_.NumberRange -eq $Range.Identity -and $_.RetentionType -match "User"}).count
                $addObject.ExcludedNumbersInRetention = @($CompleteWorkingRange | Where-Object {$_.NumberRange -eq $Range.Identity -and $_.RetentionType -match "Excluded"}).count
                $addObject.TotalNumbersInRetention = @($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $Range.Identity -and $_.InRetention -eq $True}).count
                $addObject.TotalNumbersInRetention = $addObject.TotalNumbersInRetention + $addObject.DisabledUsersInRetention
                $addObject.NumbersWithActivityInRetention = @($CompleteWorkingRange | Where-Object {$_.NumberRange -eq $Range.Identity -and $_.InRetention -eq $True -and $_.NumberActivity}).count
                $addObject.Comment = $Extension[2]
                $ReportToScreen += $addObject

        }

    #}

    Write-Host " Done!" -ForegroundColor Green
    #Report complete used and unused numbers to GridView
    if ($ReportGrid -eq $True -and $AllRanges -eq $False){$CompleteWorkingRange | Where-Object {$_.NumberRange -eq $WorkingRange.Identity} | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property DID | Out-GridView}
    if ($ReportGrid -eq $True -and $AllRanges -eq $True){$CompleteWorkingRange | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property DID | Out-GridView}

    
    #Report complete used and unused numbers to CSV
    if ($ReportCsv -eq $True -and $AllRanges -eq $False){
        $Path = "c:\_Report\"
        if(!(Test-Path -Path $path)){md $Path}
        $File = $Path+"PhoneNumbers"+(get-date -f yyyy-MM-dd)+".txt"
        $CompleteWorkingRange | Where-Object {$_.NumberRange -eq $WorkingRange.Identity} | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property DID | Export-Csv $File -noType -Encoding Unicode
        Write-Host "CSV exported to $Path" -foregroundcolor Yellow
    }
    if ($ReportCsv -eq $True -and $AllRanges -eq $True){
        $Path = "c:\_Report\"
        if(!(Test-Path -Path $path)){md $Path}
        $File = $Path+"PhoneNumbers"+(get-date -f yyyy-MM-dd)+".txt"
        $CompleteWorkingRange | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property DID | Export-Csv $File -noType -Encoding Unicode
        Write-Host "CSV exported to $Path" -foregroundcolor Yellow
    }

    #Return all info as script output, to work with outside the script
    if ($ReturnAllInfo -eq $True -and $AllRanges -eq $False){Return $CompleteWorkingRange | Where-Object {$_.NumberRange -eq $WorkingRange.Identity} | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property DID}
    if ($ReturnAllInfo -eq $True -and $AllRanges -eq $True){Return $CompleteWorkingRange | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property DID}

    else {Return $ReportToScreen}

    $ErrorActionPreference = 'Continue'

    

}

###############################################################################
## Script body
###############################################################################

#All ranges in deployment
#$AllSeries

#All numbers assigned to users
Write-Host Finding all numbers used in the deployment... -ForegroundColor yellow -NoNewline
$AllUSedNumbers = Get-AllnumbersInDeployment
#$AllUSedNumbers | Sort-Object -Property did | ft displayname, did, ext
Write-Host Done! -ForegroundColor Green

#Getting the ranges to work with based on script input
Write-Host Determining what range to work with... -ForegroundColor Yellow -NoNewline
$WorkingRange = Get-WorkingSeries 
#$WorkingRange
Write-Host  " Done!" -ForegroundColor Green

#Collecting all available numbers based on working ranges
Write-Host Finding all numbers in ranges... -ForegroundColor Yellow -NoNewline
$AvailableNumbers = Get-AllAvailableNumbers($WorkingRange)
#$AvailableNumbers
Write-Host " Done!" -ForegroundColor Green

#Creating an object that has all available and used numbers in sorted order
Write-Host Generating complete internal database of used and available numbers... -ForegroundColor Yellow -NoNewline
$CompleteWorkingRange = Get-CompleteWorkingRange $AllUSedNumbers $AvailableNumbers
#$CompleteWorkingRange | ft displayname, did, numberRange
Write-Host " Done!" -ForegroundColor Green

#Adding users that are disabled in AD, but still have a lineuri $CompleteWorkingRange  
#Uses Active Directory PowerShell module
Write-host Finding and adding users that are disabled in AD but still have a LineURI as numbers in retention... -ForegroundColor Yellow -NoNewline
$Retention = Get-Retention 
#$Retention
Write-Host " Done!" -ForegroundColor Green

#Adding numbers that are in manual retention form $NRs at the top of the script 
Write-host Finding reserved gold, silver, bronze and special numbers from internal database... -ForegroundColor Yellow -NoNewline
$GoldRetention = Get-GoldRetention 
#$GoldRetention
#Adding numbers that will be automatically classified as Gold, Silver and Bronze
if ($AutoclassifyNumbers -eq $True){
    $NumberClassification = Start-NumberClassification
}
Write-Host " Done!" -ForegroundColor Green

#Adding numbers from unassinged numbers that has Gold, Silver, Bronze or Special in their name
Write-host Finding reserved gold, silver, bronze and special numbers from Unassigned Numbers... -ForegroundColor Yellow -NoNewline
$AddedRange = Get-UnassignedRetention
#$AddedRange
Write-Host " Done!" -ForegroundColor Green

#Finding activity of all numbers within specified timeframe
if ($ReportUserActivity -eq $True){
    Write-host Connecting to Monitoring CDR database to check activity on numbers the last 30 days... -ForegroundColor Yellow -NoNewline
    $NumberActivity = Check-NumberActivity(30)
    #$NumberActivity
    Write-Host " Done!" -ForegroundColor Green
}

#Check for users that has not logged on in more than 90 days
#you need to be logged in as a user with read access to the monitoring server
Write-host Connecting to Monitoring CDR database to check for users that has not logged on for 30 days or more... -ForegroundColor Yellow -NoNewline
$CheckUserActivity = Check-USerActivity (30)
#$CheckUserActivity
Write-Host " Done!" -ForegroundColor Green

Write-host Generating report... -ForegroundColor Yellow -NoNewline
#Measuring and reporting available numbers
#Output stats to console
$Report = Get-Report $CompleteWorkingRange
if ($Report -ne $Null -and $FirstAvailable -eq $False -and $ReturnAllInfo -eq $True){Return $Report | Sort-Object -Property NumberOfAvailableNumbers -Descending }
if ($ReportPieChartHTML -eq $True){Write-host Generating Oice Charts in HTML... -ForegroundColor Yellow; Report-PieChartHTML($Report)}
elseif ($Report -ne $Null -and $FirstAvailableAsLineURI -eq $True){$LineURI= "tel:"+$Report.FirstavailableNumber+";ext="+$Report.FirstavailableExtension; Return $LineURI}
if ($Report -ne $Null -and $FirstAvailable -eq $False){$Report | Sort-Object -Property NumberOfAvailableNumbers -Descending | Select-Object Identity, NumberRangeStart, NumberRangeEnd, TotalNumbersInRange, NumberOfAvailableNumbers, FirstavailableNumber, FirstavailableExtension, TotalNumbersInRetention, GoldnumbersInRetention, SilverNumbersInRetention, BronzeNumbersInRetention, SpecialNumbersInRetention, DisabledUsersInRetention, NumbersWithActivityInRetention, Comment}
elseif ($Report -ne $Null -and $FirstAvailable -eq $True){return $Report}




