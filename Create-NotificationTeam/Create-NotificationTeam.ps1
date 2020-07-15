#Created by MVP Ståle Hansen
#Read the original post at https://msunified.net
#Make sure you are connected to Exchange Online Powershell and Microsoft Teams powershell

$Name = "M365NotificationTeam"
$Domain= "yourdomain"  #maildomain
$UserName="yourUPN"  #Teams enabled owner


#Create the Office 365 Group using the Exchange Online PowerShell Module
New-UnifiedGroup –DisplayName $Name –Alias $Name –EmailAddresses "$Name@$Domain" -owner $UserName -RequireSenderAuthenticationEnabled $False -Verbose

#This is optional, but may be a good practice initially since Office 365 Groups may clutter your Global Addressbook
Set-UnifiedGroup –Identity $Name  –HiddenFromAddressListsEnabled $true

#Create the Team using the Microsoft Teams PowerShell module
$group = New-Team -Group (Get-UnifiedGroup $Name ).ExternalDirectoryObjectId -Verbose

Get-Team -DisplayName $Name

$Group = Get-Team -DisplayName $Name
#Add Channels to the Team for Message Center
New-TeamChannel -GroupId $group.GroupId -DisplayName "Core Services Message Center" -Description "All Message Center posts related to Microsoft 365 Office Subscription, Office 365 Portal, Office for the web and other undocumented categories"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Exchange Message Center" -Description "All Message Center posts related to Exchange"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Teams and Skype Message Center" -Description "All Message Center posts related to Microsoft Teams and Skype for Business"
New-TeamChannel -GroupId $group.GroupId -DisplayName "SharePoint and OneDrive Message Center" -Description "All Message Center posts related to SharePoint and OneDrive for Business"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Yammer Message Center" -Description "All Message Center posts related to Yammer"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Intune Message Center" -Description "All Message Center posts related to Microsoft Intune"

#Add channels to the Teams for Service Health
New-TeamChannel -GroupId $group.GroupId -DisplayName "Core Services Service Health" -Description "All Health Center posts related to Microsoft 365 Office Subscription, Office 365 Portal, Office for the web and other undocumented categories"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Exchange Service Health" -Description "All Health Center posts related to Exchange"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Teams and Skype Service Health" -Description "All Health Center posts related to Microsoft Teams and Skype for Business"
New-TeamChannel -GroupId $group.GroupId -DisplayName "SharePoint and OneDrive Service Health" -Description "All Health Center posts related to SharePoint and OneDrive for Business"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Yammer Service Health" -Description "All Health Center posts related to Yammer"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Intune Service Health" -Description "All Health Center posts related to Microsoft Intune"

#Add channels to the Teams for Microsoft 365 Roadmap
New-TeamChannel -GroupId $group.GroupId -DisplayName "Core Services Roadmap" -Description "All Microsoft 365 Roadmap posts related to Microsoft 365 Office Subscription, Office 365 Portal, Office for the web and other undocumented categories"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Exchange Roadmap" -Description "All Microsoft 365 Roadmap posts related to Exchange"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Teams and Skype Roadmap" -Description "All Microsoft 365 Roadmap posts related to Microsoft Teams and Skype for Business"
New-TeamChannel -GroupId $group.GroupId -DisplayName "SharePoint and OneDrive Roadmap" -Description "All Microsoft 365 Roadmap posts related to SharePoint and OneDrive for Business"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Yammer Roadmap" -Description "All Microsoft 365 Roadmap posts related to Yammer"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Intune and Identity Roadmap" -Description "All Microsoft 365 Roadmap posts related to Microsoft Intune"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Security Roadmap" -Description "All Microsoft 365 Roadmap posts related to ATP and Cloud App Security"
New-TeamChannel -GroupId $group.GroupId -DisplayName "NextGenApps Roadmap" -Description "All Microsoft 365 Roadmap posts related to Planner, Stream, Forms, Bookings and Whiteboard"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Office Client Roadmap" -Description "All Microsoft 365 Roadmap posts related to Word, Excel, PowerPoint, Outlook, To-Do, Visio, OneNote, Project, Access"
New-TeamChannel -GroupId $group.GroupId -DisplayName "Compliance Roadmap" -Description "All Microsoft 365 Roadmap posts related to Information Protection"

#Add channels to the Teams for Office client updated
New-TeamChannel -GroupId $group.GroupId -DisplayName "Office Whats New" -Description "All Office clients updates"
