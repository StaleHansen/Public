.EXAMPLE 
   .\Set-CsCustomPresence.ps1 
   Will run the script and all of its content with the functions
 
This script will create a local custom presence XML file on your computer and set the correct registry entries so you may enjoy custom presence states even though your company has not deplyed them or you do want to use your own states.
    - Added support for Office 2016 and Skype for Business
    - Will Create XML file and add registry settings
    - Will add both english culture and local culture, add two lines with same culture is ok
    - May require elevated permissions
    - You may have to Set-ExecutionPolicy unrestricted
    - The script is not signed
Check original blogpost for the script:

https://msunified.net/2017/08/20/how-to-set-custom-presence-state-in-skype-for-business-on-your-windows-machine/