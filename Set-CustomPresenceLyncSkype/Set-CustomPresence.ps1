<#
---------------------------------------------------
Created by MVP Ståle Hansen
---------------------------------------------------

Set-CsCustomPresence.ps1 Will create a custom presence XML file and add the correct registry entries

.Notes
    - Will Create XML file and add registry settings
    - Will at both english culture and local culture, add two lines with same culture is ok
    - May require elevated permissions
    - You may have to Set-ExecutionPolicy unrestricted
    - The script is not signed

V 2.1 August 2017 - Validated the script and corrected typo
V 2.0 October 2015  - Added support for Office 2016 and Skype for Business
V 1.0 May 2014  - Initial Script

.Link
   Twitter: http://www.twitter.com/StaleHansen
   Blog: http://msunified.net
   LinkedIn: http://www.linkedin.com/in/StaleHansen
   Current Release: 
.EXAMPLE
   .\Set-CsCustomPresence.ps1
   Will run the script and all of its content with the functions
#>

$CustomPresencePath="$env:SystemDrive\_CustomPresence"
$XMLFile=$CustomPresencePath+"\CustomPresence.xml"
$xmlpathinregistry=$XMLFile
#English for native english, and there is always someone running the english Lync client that you know
$EnglishCulture=1033
#for the second line we would like to add the native language for us non-native english guys
$LocalCulture=(Get-Culture).LCID
#If you need to add more culture LCID, just add new lines below, check LCID codes here: http://msdn.microsoft.com/en-us/goglobal/bb964664.aspx

Function Create-CustomPresenceXml{
$HereString=@"
<?xml version="1.0"?>
<customStates xmlns="http://schemas.microsoft.com/09/2009/communicator/customStates">
    <customState ID="1" availability="do-not-disturb">
        <activity LCID="$EnglishCulture">Pomodoro Sprint</activity>
        <activity LCID="$LocalCulture">Pomodoro Sprint</activity>
    </customState>
    <customState ID="2" availability="Busy">
        <activity LCID="$EnglishCulture">Workshop</activity>
        <activity LCID="$LocalCulture">Workshop</activity>
    </customState>
    <customState ID="3" availability="Busy">
        <activity LCID="$EnglishCulture">Getting Coffee</activity>
        <activity LCID="$LocalCulture">Getting Coffee</activity>
    </customState>
    <customState ID="4" availability="Online">
        <activity LCID="$EnglishCulture">Skype me whenever</activity>
        <activity LCID="$LocalCulture">Skype me whenever</activity>
    </customState>
</customStates>
"@    
Set-Content -Path $XMLFile -Value $HereString
}

Function Create-RegistrySettingsOffice2013{

    if((Test-Path HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Lync) -ne $True){
        #Create the registry paths for 32 bit which also apply to 64 bit
        New-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\ -ErrorAction SilentlyContinue
        New-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\15.0\  -ErrorAction SilentlyContinue
        New-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\15.0\Lync\  -ErrorAction SilentlyContinue
    }

    if((Get-ItemProperty  HKLM:\SOFTWARE\Policies\Microsoft\Office\15.0\Lync -Name EnableSIPHighSecurityMode -ea 0).EnableSIPHighSecurityMode -eq "0") {'EnableSIPHighSecurityMode Property already exists for Office 2013'}
    else {
        Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\15.0\Lync -Name EnableSIPHighSecurityMode -Value "0" -Type DWord
        Write-Output "EnableSIPHighSecurityMode Property set for Office 2013, if it is the first time this is set, you may need to reboot your computer" 
    }

    if((Get-ItemProperty  HKLM:\SOFTWARE\Policies\Microsoft\Office\15.0\Lync -Name CustomStateURL -ea 0).CustomStateURL -eq $xmlpathinregistry) {'CustomStateURL Property already exists for Office 2013'}
    else {
        Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\15.0\Lync -Name CustomStateURL -Value $xmlpathinregistry
        Write-Output "CustomStateURL Property set for Office 2013, if it is the first time this is set, you may need to reboot your computer" 

    }

}

Function Create-RegistrySettingsOffice2016{

    if((Test-Path HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Lync) -ne $True){
        #Create the registry paths for 32 bit whcich also apply to 64 bit
        New-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\ -ErrorAction SilentlyContinue
        New-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\  -ErrorAction SilentlyContinue
        New-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Lync\  -ErrorAction SilentlyContinue
    }

    if((Get-ItemProperty  HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Lync -Name EnableSIPHighSecurityMode -ea 0).EnableSIPHighSecurityMode -eq "0") {'EnableSIPHighSecurityMode Property already exists for Office 2016'}
    else {
        Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Lync -Name EnableSIPHighSecurityMode -Value "0" -Type DWord
        Write-Output "EnableSIPHighSecurityMode Property set for Office 2016, if it is the first time this is set, you may need to reboot your computer" 
    }

    if((Get-ItemProperty  HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Lync -Name CustomStateURL -ea 0).CustomStateURL -eq $xmlpathinregistry) {'CustomStateURL Property already exists for Office 2016'}
    else {
        Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Lync -Name CustomStateURL -Value $xmlpathinregistry
        Write-Output "CustomStateURL Property set for Office 2016, if it is the first time this is set, you may need to reboot your computer" 

    }

}

#Creating folder if it does not exist
if(!(Test-Path $CustomPresencePath)){New-Item -ItemType Directory -Force -Path $CustomPresencePath}

#Running the functions
Create-CustomPresenceXml
Create-RegistrySettingsOffice2013
Create-RegistrySettingsOffice2016
