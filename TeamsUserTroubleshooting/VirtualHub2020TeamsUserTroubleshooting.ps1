break

cd c:\

cls






Start-Process "https://msunified.net/2019/07/11/my-post-migration-from-skype-to-teams-toolbox/"










$UserName = "<YourTeamsServiceAdmin>"


Import-Module MicrosoftTeams

Get-Module MicrosoftTeams

Install-Module MicrosoftTeams -Force















#Connect to SFBO via Teams
Connect-MicrosoftTeams -AccountId $UserName
$SkypeSession = New-CsOnlineSession
Import-PSSession $SkypeSession -AllowClobber 
C:\Temp\.\Enable-CsOnlineSessionForReconnection.ps1






#individualcheck
$user="<TeamsUser>"
Get-CsOnlineUser $User | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, `
EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, DialPlan, TenantDialPlan, OnlineVoiceRoutingPolicy, `
LineURI, OnPremLineURI, OnlineDialinConferencingPolicy, TeamsVideoInteropServicePolicy, TeamsCallingPolicy, HostingProvider, `
InterpretedUserType, VoicePolicy,CountryOrRegionDisplayName

Start-Process "https://msunified.net/2019/07/11/my-post-migration-from-skype-to-teams-toolbox/"










$MigrationBatch = "VirtualHub"

#create a variable with the import of users to migrate and count them
$Users = import-csv -Path "C:\Temp\$MigrationBatch.csv" -Delimiter ";" -Encoding UTF7
$UsersCount = $users.count
$UsersCount


#Simple Bulk
foreach ($User in $Users.email){
Get-CsOnlineUser $User | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, `
EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, DialPlan, TenantDialPlan, OnlineVoiceRoutingPolicy, `
LineURI, OnPremLineURI, OnlineDialinConferencingPolicy, TeamsVideoInteropServicePolicy, TeamsCallingPolicy, HostingProvider, `
InterpretedUserType, VoicePolicy,CountryOrRegionDisplayName 
}




#Validation Batch
$migrateduseroutput=@()
[int]$count=$Users.count
[int]$allusers=$Users.count
$i=0
foreach ($User in $Users.email){
    $migrateduseroutput += (Get-CsOnlineUser $User | select-object UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, `
    EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, DialPlan, TenantDialPlan, OnlineVoiceRoutingPolicy, `
    LineURI, OnPremLineURI, OnlineDialinConferencingPolicy, TeamsVideoInteropServicePolicy, TeamsCallingPolicy, HostingProvider, `
    InterpretedUserType, VoicePolicy,CountryOrRegionDisplayName)
    Start-Sleep -Milliseconds 400
    if($i -eq 100 -or $count -eq $allusers){
        # Writing feedback
        Write-Host "$count users processed of $allusers" -f yellow
        # Counter is being reset
        $i = 0
    }
    $i++
    $count--
}



#$migrateduseroutput.count
$migrateduseroutput | Export-Csv -Path C:\Temp\PreCheckBatch8Part1.csv -Delimiter ";" -NoTypeInformation -encoding UTF8






#Validate UPN
Get-CsOnlineUser $User | Format-List UserPrincipalName, DisplayName, SipAddress











#validate Enabled
Get-CsOnlineUser $User | Format-List Enabled,TeamsUpgradeEffectiveMode, EnterpriseVoiceEnabled, OnlineDialinConferencingPolicy








#Validate Licenses

Import-Module AzureADPreview

Connect-AzureAD -AccountId $UserName

#Find all licenses
Get-AzureADSubscribedSku | Select SkuPartNumber,SkuID

Start-Process "https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference"


$ENTERPRISEPACK="6fd2c87f-b296-42f0-b197-1e91e994b900"
$DESKLESSPACK="24b585984-651b-448a-9e53-3b10f069cf7f"
$SPE_E5="06ebc4ee-1bb5-47dd-8120-11324bc54e06"

$MCOPSTN_5="11dee6af-eca8-419f-8061-6864517c1875"





#Check Assigned licenses
Get-AzureADUser -ObjectID $user | Select -ExpandProperty AssignedLicenses












#Check for non licensed users
$UnlicensedUsers=@()
$NOADUsers=@()
$ENTERPRISEPACK="6fd2c87f-b296-42f0-b197-1e91e994b900"
$DESKLESSPACK="24b585984-651b-448a-9e53-3b10f069cf7f"
$SPE_E5="06ebc4ee-1bb5-47dd-8120-11324bc54e06"
foreach ($userUPN in $Users.email){
    $userList = Get-AzureADUser -ObjectID $userUPN | Select -ExpandProperty AssignedLicenses
    if ($userList -eq $null){write-host $userUPN "does not exist in AAD";$NOADUsers+=$userUPN}
    elseif ($userList.SkuId -notcontains $ENTERPRISEPACK -and $userList.SkuId -notcontains $DESKLESSPACK -and $userList.SkuId -notcontains $SPE_E5){write-host $userUPN "not licensed";$UnlicensedUsers+=$userUPN}
    $userlist=$null
}

$UnlicensedUsers | Out-File -file "C:\Temp\UnlicensedTeamsFest.csv"
$NOADUsers | Out-File -file "C:\Temp\NOADTeamsFest.csv"









#validate Calling and location
Get-CsOnlineUser $User | Format-List UserPrincipalName, City, UsageLocation, DialPlan, TenantDialPlan, OnlineVoiceRoutingPolicy,LineURI, OnPremLineURI, TeamsCallingPolicy, VoicePolicy,CountryOrRegionDisplayName,Country
Get-CsOnlineUser "<AnotherTeamsUser>"| Format-List UserPrincipalName, City, UsageLocation, DialPlan, TenantDialPlan, OnlineVoiceRoutingPolicy,LineURI, OnPremLineURI, TeamsCallingPolicy, VoicePolicy,CountryOrRegionDisplayName,Country












#TenantDialplan
Start-Process "https://www.ucdialplans.com/"


#Creating batch for assigning onlinevoiceroutingpolicy
$Batchname = Get-Date -Format "MM-dd-yyyy-HH-mm"
New-CsBatchPolicyAssignmentOperation -PolicyType OnlineVoiceRoutingPolicy -PolicyName "Unrestricted” -Identity $Users.email -OperationName $Batchname
$Users.email | Out-File -FilePath "C:\temp\$Batchname.csv" 

Get-CsBatchPolicyAssignmentOperation 









#Validate location, look at usagelocation, city, country and lineuri
Get-AzureADUser | Where-Object {$_.UserPrincipalName -match $User} | Select-Object -Property UserPrincipalName,UsageLocation

Set-AzureADUser -ObjectId $objID -UsageLocation US -Verbose

#validate Calling and location
Get-CsOnlineUser $User | Format-List UserPrincipalName, City, UsageLocation, DialPlan, TenantDialPlan, OnlineVoiceRoutingPolicy,LineURI, OnPremLineURI, TeamsCallingPolicy, VoicePolicy,CountryOrRegionDisplayName

#Change Audio Conferencing number for existing meetings
Get-CsMeetingMigrationStatus -SummaryOnly
Start-CsExMeetingMigration -Identity my.buddy@cloudway.no








#Hybrid identity and Skype for Business
Get-CsOnlineUser $User | Format-List UserPrincipalName,HostingProvider,InterpretedUserType 

Start-Process "https://get-itips.capazero.net/posts/sfbonprem-interpretedusertype"




#Active Directory On-Premises command
Get-ADUser  -Filter "UserPrincipalName -eq '$user'" -property *| Set-ADUser –clear 'msRTCSIP-DeploymentLocator'





#BulkChange On-Premises
foreach ($User in $users) {
 
 
    $u=$user.UserPrincipalName
    #Get all msRTCSIP properties for a user that has a value
    $Properties = Get-ADUser -Filter {UserPrincipalName -eq $u} -Properties * | Select-Object -Property 'msRTCSIP*'
    $LineUri = Get-ADUser -Filter {UserPrincipalName -eq $u} -Properties * | Select-Object -Property UserPrincipalName,'msRTCSIP-Line'

    if ($properties -match "srv"){
        #Clear all properties for a user
        Get-ADUser -Filter {UserPrincipalName -eq $u} -Properties * | Set-ADUser -clear ($Properties | Get-Member -MemberType "NoteProperty " | % { $_.Name })
        $properties
 
    }
}









#set lineuri if Direct Routing
Set-CsUser $user -OnPremLineURI "tel:+19175428572" -EnterpriseVoiceEnabled $True

#If unable to set onpremlineuri in Office 365, the msRTCSIP-Line is probably populated on-premises
Get-ADUser -Filter "UserPrincipalName -eq '$user'" -property * | Set-ADUser –clear 'msRTCSIP-Line'

#It is ok to work with numbers on-premises
Get-ADUser  -Filter "UserPrincipalName -eq '$User'" -property * | Set-ADUser –replace @{'msRTCSIP-Line'="tel:+19175428572"}

#Set lineuri if Calling Plan
Set-CsUser $user -LineURI "tel:+19175428572" -EnterpriseVoiceEnabled $True














#Simple Bulk
foreach ($User in $Users.email){
Get-CsOnlineUser $User | Format-List UserPrincipalName, DisplayName, SipAddress, Enabled, TeamsUpgradeEffectiveMode, `
EnterpriseVoiceEnabled, HostedVoiceMail, City, UsageLocation, DialPlan, TenantDialPlan, OnlineVoiceRoutingPolicy, `
LineURI, OnPremLineURI, OnlineDialinConferencingPolicy, TeamsVideoInteropServicePolicy, TeamsCallingPolicy, HostingProvider, `
InterpretedUserType, VoicePolicy,CountryOrRegionDisplayName 
}

