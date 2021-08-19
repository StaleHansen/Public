#This script is authored by MVP Ståle Hansen, @StaleHansen on Twitter, feedback can be provided here with more info available: https://msunified.net/2021/08/18/find-available-phone-numbers-with-get-teamsnumbers-ps1/
<#
.SYNOPSIS
This script will get the next available number of any provided number range from Unassigned Numbers, Array or input to the script and will generate a full report and a summary per range

.NOTES
V1.0 - Initial version by 

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


.EXAMPLE
#Return all avaialble numbers for a number range
Rune the script in PowerShell ISE Window. After the script has run, then type $Report[0].AllAvailableNumbers where 0 is the identity of the number range

.INPUTS
The script does not support piped input at this time

.OUTPUTS
The script produces fully functional powershell output either you want the summary or use the $ReturnAllInfo $True switch to return not just a summary but all info as seen in the GridView

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
   $ReportCSV=$True,
   $ReportHTML=$False,
   $ReportEmail=$False,
   $ReportGrid=$True,
   $ReportUserActivity=$False,
   $AutoclassifyNumbers=$True,
   $ReserveBronzeNumbers=$False,
   $ReportScreen=$True

)

#Global Variables
$AvailableNumbers=$Null




########################################################
##             Phone Number Range Database
##          Add your custom ranges in advance
########################################################
#Here you can add all your custom number ranges
$CRs = @()
$CRs += ,@("NorwayOperatorConnect4050","+4764974050","+4764974059")
$CRs += ,@("SwedenCallingPlan4050","+46850241520","+46850241540")
$CRs += ,@("NorwayDirectRouting8050","+4764978050","+4764979999")


########################################################
##           Phone Number Retention Database
##      For your Gold, silver and special numbers
########################################################
#Here you can add all your custom number ranges
$NRs = @()
$NRs += ,@("Reserved for Vanity","+4764974050")

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
    $userUris = Get-CsOnlineUser -Filter {OnPremLineURI -ne $Null} -WarningAction SilentlyContinue | % { NewLyncNumberFromAdContact "User" $_ }
    #$userUris.count

    # combine all results together
    $allUsedNumbers = New-Object System.Collections.ArrayList 
    foreach($list in @($userUris,$plUris,$analogUris,$caUris,$rgsUris,$dialinUris,$exumUris,$tepUris,$MTRUris)) {
	    if($list -and $list.Length -gt 0) {
		    $allUsedNumbers.AddRange($list)
	    }
    }

    Return $allUsedNumbers
}


###############################################################################
## Finding the range we want to work with
###############################################################################

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
## Create an object that has all users, used and available numbers
###############################################################################

function Add-CompleteWorkingRange{

    New-Object -TypeName PSCustomObject -Property @{
        Type = $null
        LineURI = $null
        DisplayName = $null
        SipAddress = $null
        Identity = $null
        Name = $null
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
        AllAvailableNumbers = $null

    }

}

function Get-CompleteWorkingRange($FunctionAllUSed, $FunctionAvailable){

    #Find all users with a number within a number range
    $MeasureObject = @()

    If($FunctionAvailable){
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
## Measuring and reporting available numbers
###############################################################################

Function Add-Report {
    New-Object -TypeName PSCustomObject -Property @{
        Name = $Null
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
        AllAvailableNumbers = $Null
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
    if ($ReportScreen -eq $True){
        $ReportToScreen = @()
        $addObject = @()
        if($AllRanges -eq $True){$Tempseries = $AllSeries; $Tempseries += $ReservedSeries; $WorkingRange = $Tempseries}

        $i=0
        foreach ($Range in ($WorkingRange | Where-Object {$_.identity -ne $Null})){
                #write-host $Range.Identity
                $Extension=$Null
                $addObject = Add-Report
                $addObject.Identity = $i
                $addObject.Name = $Range.Identity
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
                $addObject.Comment = $Comment
                $addObject.AllAvailableNumbers = ($CompleteWorkingRange | Where-Object {$_.displayname -eq $Null -and $_.NumberRange -eq $Range.Identity -and $_.InRetention -eq $False}).did
                $ReportToScreen += $addObject

                $i++
        }

    }

    Write-Host " Done!" -ForegroundColor Green
    #Report complete used and unused numbers to GridView
    if ($ReportGrid -eq $True -and $AllRanges -eq $False){$CompleteWorkingRange | Where-Object {$_.NumberRange -eq $WorkingRange.Identity} | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property DID | Out-GridView}
    if ($ReportGrid -eq $True -and $AllRanges -eq $True){$CompleteWorkingRange | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property DID | Out-GridView}

    
    #Report complete used and unused numbers to CSV
    if ($ReportCsv -eq $True -and $AllRanges -eq $False){
        $Path = "c:\_Report\"
        if(!(Test-Path -Path $path)){md $Path}
        $File = $Path+"PhoneNumbers"+(get-date -f yyyy-MM-dd)+".csv"
        $CompleteWorkingRange | Where-Object {$_.NumberRange -eq $WorkingRange.Identity} | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property NumberRange -Descending | Export-Csv $File -noType -Encoding Unicode -Delimiter ";"
        Write-Host "CSV exported to $Path" -foregroundcolor Yellow
    }
    if ($ReportCsv -eq $True -and $AllRanges -eq $True){
        $Path = "c:\_Report\"
        if(!(Test-Path -Path $path)){md $Path}
        $File = $Path+"PhoneNumbers"+(get-date -f yyyy-MM-dd)+".csv"
        $CompleteWorkingRange | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,Comment | Sort-Object -Property NumberRange -Descending | Export-Csv $File -noType -Encoding Unicode -Delimiter ","
        Write-Host "CSV exported to $Path" -foregroundcolor Yellow
    }

    #Return all info as script output, to work with outside the script
    if ($ReturnAllInfo -eq $True -and $AllRanges -eq $False){Return $CompleteWorkingRange | Where-Object {$_.NumberRange -eq $WorkingRange.Identity} | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,AllAvailableNumbers,Comment | Sort-Object -Property DID}
    if ($ReturnAllInfo -eq $True -and $AllRanges -eq $True){Return $CompleteWorkingRange | Select-Object NumberRange,DisplayName,SipAddress,Type,LineURI,DID,ext,VoicePolicy,InRetention,RetentionType,OfflineMoreThan30Days,NumberActivity,AllAvailableNumbers,Comment | Sort-Object -Property DID}

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

#Adding numbers that are in manual retention form $NRs at the top of the script 
Write-host Finding reserved gold, silver, bronze and special numbers from internal database... -ForegroundColor Yellow -NoNewline
$GoldRetention = Get-GoldRetention 
#$GoldRetention
#Adding numbers that will be automatically classified as Gold, Silver and Bronze
if ($AutoclassifyNumbers -eq $True){$NumberClassification = Start-NumberClassification}
Write-Host " Done!" -ForegroundColor Green

Write-host Generating report... -ForegroundColor Yellow -NoNewline
#Measuring and reporting available numbers
#Output stats to console

$Report = Get-Report $CompleteWorkingRange
if ($Report -ne $Null -and $FirstAvailable -eq $False -and $ReturnAllInfo -eq $True){Return $Report | Sort-Object -Property NumberOfAvailableNumbers -Descending }
elseif ($Report -ne $Null -and $FirstAvailableAsLineURI -eq $True){$LineURI= "tel:"+$Report.FirstavailableNumber+";ext="+$Report.FirstavailableExtension; Return $LineURI}
if ($Report -ne $Null -and $FirstAvailable -eq $False){$Report | Sort-Object -Property NumberOfAvailableNumbers -Descending | Select-Object Identity, Name, NumberRangeStart, NumberRangeEnd, TotalNumbersInRange, NumberOfAvailableNumbers, FirstavailableNumber, FirstavailableExtension, TotalNumbersInRetention, GoldnumbersInRetention, SilverNumbersInRetention, BronzeNumbersInRetention, SpecialNumbersInRetention, DisabledUsersInRetention, NumbersWithActivityInRetention,AllAvailableNumbers,Comment}
elseif ($Report -ne $Null -and $FirstAvailable -eq $True){return $Report}
if ($ReportScreen -eq $True) {$ReportToScreen |  Select-Object -Property Identity,Name,NumberRangeStart,NumberRangeEnd,TotalNumbersInRange,NumberOfAvailableNumbers,FirstavailableNumber,FirstavailableExtension,GoldnumbersInRetention,SilverNumbersInRetention,BronzeNumbersInRetention,SpecialNumbersInRetention,ExcludedNumbersInRetention,DisabledUsersInRetention,TotalNumbersInRetention,AllAvailableNumbers,Comment}

