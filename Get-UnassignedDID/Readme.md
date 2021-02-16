Note: a new script has been created called Get-SfBNumbers.ps1. It finds available numbers, activity on numbers and users that have numbers but are disabled. It can export to gridview, csv and it even creates pie charts. Check it out at http://sfbnumbers.net/

Script for Lync Server 2010 and 2013 that searches through all users, services and devices to find all used numbers and uses that to find unused numbers in a given number serie for Lync. The numberserie is based on Unassigned Numbers voice feature in Lync and the script supports numbers larger than 10 digits.

Features:

- New feature as of 13.05.2014: you are now able to use a custom number range
    - See original blogpost: http://lyncnumbers.net/2014/05/13/use-a-custom-number-range-with-get-unusednumbers-ps1/
- The entire script is a function and is ready to be reused in scripts to enable users for Enteprise Voice 
- The function can take a Unassinged Number Identity and return the value of a new unassigned number
      - If you don’t provide input to the function it will go through all Unassigned Numbers in the deployment
      - It will check if there is no Unassigned Number range defined, and point you to TechNet for how to configure it
- Support for numbers larger than 10 digits added
     - A problem with the old script was that it did not support numbers with more than 10 digits (Int32)
     - I have rewritten the code to create a list of the number range that supports Int64 values which should be enough digits for phone numbers
- I recommend working with the numbers in full E.164 format to be able to support numbers with different countrycode length
- Separate function for finding and creating lineuri string with extension is available
     - It will check if extension exist already in the deployment and write a warning
- Searching through the numbers is done using a separate function which makes it easy to reuse in your own routines

Who should use it?
Administrators looking for simplifying administration of numbers assigned to users and devices in a Lync deployment
Lync PRO's when advising and deploying Enterprise Voice

Planned updates:

- Export all available numbers to file, grid view or xml
- Create parameters for running the script as a script and not just funcions
- Sign the script  

