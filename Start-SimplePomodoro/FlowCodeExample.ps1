[int]$Minutes = 25 #Duration of your Pomodoro Session, default is 25 minutes
[string]$Secret = "MySecret" #Secret for the flow trigger
[string]$AutomateURI = "YourFlowTriggerURI" #The URI used in the webrequest to your flow


#Invoking PowerAutomate to change set current time on your Focus time calendar event, default length is 25 minutes
    $body = @()
    $body = @"
        { 
            "Duration":$Minutes,
            "Secret":"$Secret"
        }
"@
Invoke-RestMethod -Method Post -Body $Body -Uri $AutomateURI -ContentType "application/json"