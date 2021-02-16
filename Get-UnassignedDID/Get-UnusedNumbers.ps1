<#
---------------------------------------------------
Created by MVP Ståle Hansen
---------------------------------------------------

Get-UnusedNumbers.ps1 is one of the lyncnumbers.net project. Please visit our website for more information on the script, and updates

.Notes
    - This script is for both Enterprise and Standard Edition Server.
    - The script is buildt around functions that you can reuse for your own rutines
    - The script checks for and uses Unassigned Numbers and must be defined in E.164 format
    - The script should be able to run without any modification.
    - The script must be run on a server where Lync PS is available.
    - The script must be run in a PS3 environment, to load all nessecary modules automatically
    - I highly recommend you test the script in your Lab, before running in your production environment
    - You can use custom numberseries and use unassigned numbers as input to the function

V 1.0 January 2011  - Initial Script
                    - Gets numbers from users, Exchange UM Contacts and Dial In Conferencing
                    - Lists total available numbers, total numbers, and what numbers are available
                    - Fixed country code to +47
V 1.2 February 2011 - Added TrustedApplicationEndpoint filter as well
                    - Added supression of error messages when removing “tel:+47″ from arrays. Got errors if they where empty
                    - Added finding numbers for Response Groups, thanks to Paul-Christiaan Diks
V 1.6 April 2011    - After some great feedback I have updated the script to be much more dynamic by supporting different country codes in the same deployment
                    - The only customization (in theory) you have to do in the script is changing the length of your country code ◾This can be done in line 019
                    - Added check if Unassigned Numbers are in E.164 format, if its not, continue to the next number serie
                    - Added finding common area phones with line uri, thanks to colleague Jarle Utne
                    - Added finding numbers for users enabled with private lines, thanks again to Paul-Christiaan Diks
                    - Added line $used=$used | ForEach-Object {$_.ToLower()} to convert the $used array to all lower case letters, because the {$_.Replace(“tel:+47″, “”) would not work on uppercase letters
                    - Added finding numbers for analog devices, thanks to Marjus Sirvinsks for the tip
V 2.0 April 2013    - The entire script is a function and is ready to be reused in scripts to enable users for Enteprise Voice 
                    - The function can take a Unassinged Number Identity and return the value of a new unassigned number
                    - If you don’t provide input to the function it will go through all Unassigned Numbers in the deployment
                    - It will check if there is no Unassigned Number range defined, and point you to TechNet for how to configure it
                    - Support for numbers larger than 10 digits added 
                    - A problem with the old script was that it did not support numbers with more than 10 digits (Int32)
                    - I have rewritten the code to create a list of the number range that supports Int64 values which should be enough digits for phone numbers
                    - I recommend working with the numbers in full E.164 format to be able to support numbers with different countrycode length
V 3.0 February 2014 - Created a new function for getting all numbers in the deployment, now also including Get-CsMeetingRoom
                    - This function is easy to reuse in your own routines
                    - Added support for extensions with the function  Get-Extension -Lineuri (Get-Unused "<your unusued number serie name>") -ExtLength 4
                    - Get-Extension will check if the extension is in use and give a warning if there is duplicate extensions in the deployment
V 4.0 May 2014      - Added ability to use custom ranges not defined by unused numbers to the Get-Unused function, see the last example


.Link
   Lyncnumbers.net: http://lyncnumbers.net
   Twitter: http://www.twitter.com/StaleHansen
   Blog: http://msunified.net
   LinkedIn: http://www.linkedin.com/in/StaleHansen
   Current Release: 
.EXAMPLE
   .\Get-UnusedNumbers.ps1
   Will run the script and all of its content with the functions 
.EXAMPLE
    cls
    $FirstNumber=Get-Unused "<your unusued number serie name>"
    Write-Host #to give some air in the output
    $FirstNumber #outputs the return value from the function    
    Write-Host #to give some air in the output
    $NumberWithExt=Get-Extension -Lineuri $FirstNumber -ExtLength 4
    $NumberWithExt
    Write-Host #to give some air in the output
    $FinalLineUri="tel:+"+$NumberWithExt
    $FinalLineUri
    Write-Host #to give some air in the output
    Write this within the script and you will use the functions within
    Will give you the first next available number and will generate and extension in the correct format
.EXAMPLE
    cls
    Get-Unused
    Will return all unassigned number series with all available numbers
.EXAMPLE
    cls
    Get-Unused -RangeStart +47232323001111 -RangeEnd +47232323001190 -Name "Custom Range Norway"
    Will return next available number for the range you specify
.EXAMPLE
    cls
    Get-Unused -RangeStart +47232323001111 -RangeEnd +47232323001190 -Name "Custom Range Norway" -ListAll $true
    Will return all numbers for the specified custom number range
#>



[System.Console]::ForegroundColor = [System.ConsoleColor]::White
clear-host
 
Write-Host "Script for finding unused numbers in Lync Server 2010/2013, by MVP Ståle Hansen"
Write-Host
 
Function Get-Unused{
     param(
    $Unassigned=$Null,
    $RangeStart=$Null,
    $RangeEnd,
    $Name="CustomRange",
    $ListAll=$false
    )
  

    if($RangeStart -ne $Null){
            if($ListAll -eq $false){
                $UnassingedRun = New-Object -TypeName PSCustomObject -Property @{
                    Identity = $Name
                    NumberRangeStart = "tel:"+$RangeStart
                    NumberRangeEnd = "tel:"+$RangeEnd
                }
                $Unassigned="NotNull"
            }
            else{
                $UnassingedRun = New-Object -TypeName PSCustomObject -Property @{
                    Identity = $Name
                    NumberRangeStart = "tel:"+$RangeStart
                    NumberRangeEnd = "tel:"+$RangeEnd
                }
            }
    }  
    elseif($Unassigned -ne $Null) {$UnassingedRun=(Get-CsUnassignedNumber $Unassigned)}
    elseif($Unassigned -eq $Null){
        if((Get-CsUnassignedNumber) -eq $Null){
            Write-Host "You do not have any Unassigned Numbers defined"
            Write-Host
            Write-Host "Go to this TechNet article to see how-to:"
            Write-Host "http://technet.microsoft.com/en-us/library/gg412748.aspx"
            Write-Host
            Write-Host "Also see how to configure the announcement service:"
            Write-Host "http://technet.microsoft.com/en-us/library/gg412783.aspx"
            Write-Host
            [System.Console]::ForegroundColor = [System.ConsoleColor]::Gray
 
        }

        else {$UnassingedRun=(Get-CsUnassignedNumber)}
    
    }
 
    foreach ($Serie in ($UnassingedRun)) {
 
    #The CountryCodeLength is the length of you countrycode. I recommend to leave it at zero and list the numbers as fully E.164 format.
    #If you want to remove more than 2 digits, change the $CountryCodeLength
    $CountryCodeLength=0
    #The "tel:+" string is the +5 lenght that is added in the next line
    $CountryCodeLength=$CountryCodeLength+5
 
    #Now we get the replace string so that all numbers can be converted to an int
    #In the norwegian case this value becomes tel:+47
    $ReplaceValue=($Serie.NumberRangeStart).Substring(0,$CountryCodeLength)
 
    #Check to see if Unassigned Numbers are in E.164 format, if its not, continue to the next number serie
    if (($ReplaceValue.Substring(0,5)) -ne "tel:+"){
        Write-Host "The script requires that Unassigned Numbers are populated in E.164 format" -Foregroundcolor Yellow
        Write-Host "It appears that the number range " -nonewline
        Write-Host $Serie.Identity -nonewline -Foregroundcolor Green
        Write-Host " is not in this format"
        Write-Host
        Continue
    }
 
    #To see what your $ReplaceValue is, untag the next line
    #Write-Host Value to be replaced is $ReplaceValue
 
    $NumberStart=$Serie.NumberRangeStart | ForEach-Object {$_.Replace($ReplaceValue, "")}
    $NumberEnd=$Serie.NumberRangeEnd | ForEach-Object {$_.Replace($ReplaceValue, "")}
 
    #Convert the range to a Int64 to be able to manager numbers with more than 10 digits
    $NumberStartInt64=[System.Convert]::ToInt64($NumberStart)
    $NumberEndInt64=[System.Convert]::ToInt64($NumberEnd)
 
    $Ser=$Null
    $Ser= New-Object System.Collections.Arraylist
    [Void]$Ser.Add($NumberStartInt64)
    #$Ser.gettype()
    $Value=$NumberStartInt64+1
    #$Value
    while ($value -lt $NumberEndInt64){
        [Void]$Ser.Add($value)
        $value++
    }
 
    [Void]$Ser.Add($value)
    #Write-Host $Ser
 
    #Get all the numbers used in the solution regardless of number range, removing existing extensions
   
    $Used=Get-AllnumbersInDeployment
    $Used=$Used | ForEach-Object {$_.Replace($ReplaceValue, "")}    
    $Used=$Used | ForEach-Object {$_.split(';')[0]}
 
    
    #Find all the numbers that are in use and part of the unassigned number serie
    $AllUsed=@()
    foreach($Series in $Ser){foreach($UsedNumber in $Used){
        if($Series -eq $UsedNumber){$AllUsed+=$UsedNumber}
        }
    }
 
    #Find all the numbers that are not in use
    $ListUnUsed=@()
    $ComparisonResult=compare-object $Ser $AllUsed
    foreach($UnUsed in $ComparisonResult){
        if($UnUsed.SideIndicator -eq '<='){$ListUnUsed+=$UnUsed.InputObject;$FreeSize++}
    }
 
    #Find how many free numbers there are in the range
    $RangeSize=($NumberEndInt64-$NumberStartInt64)+1
    $TotalUsed = $RangeSize-$FreeSize
    $TotalFree = $RangeSize-$TotalUsed
 
    Write-Host "Total free numbers in number range " -nonewline
    Write-Host $Serie.Identity -NoNewLine -Foregroundcolor Green
    Write-Host ", " -NoNewLine
    Write-Host $TotalFree -NoNewLine
    Write-Host " of"$RangeSize
    Write-Host "This range starts with " -NoNewLine
    Write-Host +$NumberStart -NoNewLine -Foregroundcolor Green
    Write-Host " and ends with " -NoNewLine
    Write-Host +$NumberEnd -Foregroundcolor Green
 
    $FreeSize=$NULL
 
    #Lists all the unused numbers if L is pressed or just return the list if a Unassigned number serie is specified for the function
    if ($Unassigned -ne $Null){Return $ListUnUsed[1];break}
    else {
            Write-Host "To list available numbers, press "-NoNewLine
            Write-Host "L" -NoNewLine -Foregroundcolor Green
            $opt = Read-Host " else press Enter"
            if($opt -eq "L"){Return $ListUnUsed}
   }
 
    Write-Host
    $ListUnUsed=$NULL
    $UsedNumbers=$NULL
    $TotalFree=$NULL
    }
 
    [System.Console]::ForegroundColor = [System.ConsoleColor]::Gray
    $Unassigned=$Null
 
}

Function Get-AllnumbersInDeployment{

    $ErrorActionPreference = 'SilentlyContinue'
 
    $Used=Get-CsUser -Filter {LineURI -ne $Null} | Select-Object LineURI | out-string -stream
    $Used+=Get-CsUser -Filter {PrivateLine -ne $Null} | Select-Object PrivateLine | out-string -stream
    $Used+=Get-CsAnalogDevice -Filter {LineURI -ne $Null} | Select-Object LineURI | out-string -stream
    $Used+=Get-CsCommonAreaPhone -Filter {LineURI -ne $Null} | Select-Object LineURI | out-string -stream
    $Used+=Get-CsExUmContact -Filter {LineURI -ne $Null} | Select-Object LineURI | out-string -stream
    $Used+=Get-CsDialInConferencingAccessNumber -Filter {LineURI -ne $Null} | Select-Object LineURI | out-string -stream
    $Used+=Get-CsTrustedApplicationEndpoint -Filter {LineURI -ne $Null} | Select-Object LineURI | out-string -stream
    $Used+=Get-CsRgsWorkflow | Select-Object LineURI | out-string -stream
    $Used+=Get-CsMeetingRoom | Select-Object LineURI | out-string -stream
    $Used=$Used | ForEach-Object {$_.ToLower()}
     
    $ErrorActionPreference = 'Continue'

    Return $Used
}



Function Get-Extension{
 
    param(
    [Parameter(Mandatory=$True)][string]$Lineuri,
    [Parameter(Mandatory=$False)][int]$ExtLength=4
    )
    #Get the extension with the given length
    $Length=$Lineuri.Length
    $RemoveLength=($Length-$ExtLength)
    $Extension=$lineuri.Remove(0,$RemoveLength)
    
    #Check if Extension exist, if it does, add one digit
    $Used=Get-AllnumbersInDeployment
    $Used=$Used | ForEach-Object {$_.split('=')[-1]}

    #Check if the extension is in use
    foreach($UsedExt in $Used){
        if($UsedExt -match $Extension){
            Write-Warning "found duplicate extension $Extension, make sure to use unique pin"
            break
        }
    }
    
    #Outputting the new LineURI with extension
    $NewLineUri=$Lineuri+";ext="+$Extension
    Return $NewLineUri
} 

#Edit this section for finding unused numbers in your deployment

#Get-Unused

#Get-Unused -RangeStart +47232323001111 -RangeEnd +47232323001190 -Name "Custom Range Norway" -ListAll $true

#section below will find next available number in the unassaigned numbers serie your secify and will also find the extension

cls
$FirstNumber=Get-Unused -RangeStart +47232323001111 -RangeEnd +47232323001190 -Name "Custom Range Norway" #-ListAll $true
Read-Host "Press enter to continue" #to give some air in the output
$FirstNumber #outputs the return value from the function    
Read-Host "Press enter to continue" #to give some air in the output
$NumberWithExt=Get-Extension -Lineuri $FirstNumber -ExtLength 4
$NumberWithExt
Read-Host "Press enter to continue" #to give some air in the output
$FinalLineUri="tel:+"+$NumberWithExt
$FinalLineUri
Read-Host "Press enter to continue" #to give some air in the output
