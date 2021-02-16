 .Notes

This script works for Lync Server 2010 and Lync Server 2013
Run Lync Management Shell in Administrative mode
Use the Function Get-RemoteProgram by Jaap Brasser, modified by MVP Jan Egil Ring
If you can not connect to a server make sure Remote Registry Service is running on the machine you are unable to reach
Port 139 and 445 needs to be open on the server you are trying to scan
Tested OK in 2010 and 2013 environments
The script will be updated as new CU’s will arrive, but will also flag unknown CU versions and let you download the latest one
V1.0 May 2014 – Initial Script

Scans registry on all servers in a Lync deployment to find current installed version
You shall not use WMI to query MSI packets as that will invoke a reconfiguraton
Read more here: http://powershell.org/wp/forums/topic/alternatives-to-win32_product/
Checks against a database in this script to find the correct CU level
Will ask if you want to download the current CU if there are servers which is not on the latest patch level
Using function New-FileDownload by MVP Pat Richard
Will open the download location after the file is downloaded
No Parameters are required
V1.1 August 2014 - Added Cumulative Update August 2014 (CU5)
V1.2 September 2014 - Added Lync Server 2013 Cumulative Update (CU6)
V1.3 November and December updates added

 