As shown at Microsoft Ignite 2015 here is the script demoed called Get-SfBNumbers.net. At this time the tool is in beta since it has only been tested in a handfull of deployments. More examples and discussion can be found at http://SfBNumbers.net

The script support Lync Server 2010, Lync Server 2013 and Skype for Business Server 2015

Description
Updated to version 1.0 with new features, fixes and optimizations

Added and verified $ReportPieChartHTML parameter and set it to $True as default
Fixed bug when connecting to SQL monitoring CDR databases with named instances and mirrored SQL
Optimized adding user activity to internal database and added parameter $ReportUserActivity that is default set to $False that will skip it since it may take up to 20 minutes
Added and verified options to automatically classify numbers to Gold and Silver with parameter $AutoclassifyNumbers that is default set to $True
Thanks to Paul Valiant for his script
Added option to classify Bronze numbers with parameter $ReserveBronzeNumbers which is default set to $False, do not run
Thanks to Paul Valiant for his script
Added check if Server 2010 to skip Get-CsMeetingRooms as the cmdlet does not exist in 2010
Added option to specify extension length with parameter $ExtensionLength that is default set to 4
This script will get the next available number of any provided number range from

Unassinged Numbers
array in the script
parameter input when running the script
It will check for

disabled users in Active Directory
connect to the LcsCDR to get users that has not logged on for 30 days or more
Connect to LcsCDR to check for activity on numbers, both unassigned and assigned
It supports

Lync Server 2010, Lync Server 2013, Skype for Business Server 2015
Reservation of numbers based on categories like Gold, Silver, Bronze, Special and Excluded
From Array in the script or Unassigned Numbers
Reservations of number ranges within larger number ranges for special services like future callcenter or response groups
Retention of numbers based on disabled in Active Directory, last logon time and activity on numbers
Reporting supported

Default behaviour is GridView with all numbers and summary in the PowerShell window
Export to of all information to CSV
Create a summary in PieChart as HTML