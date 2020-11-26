Function Start-SimplePomodoro {

<#
      .SYNOPSIS
      Start-SimplePomodoro is a function command to start a new Pomodoro session with additional actions. This is a simplified version of the Start-Pomodoro 
      .DESCRIPTION

        By MVP Ståle Hansen (http://msunified.net) with modifications by Jan Egil Ring (https://github.com/janegilring)
        Pomodoro function by Nathan.Run() http://nathanhoneycutt.net/blog/a-pomodoro-timer-in-powershell/
        Note: for desktops you need to enable presentation settings in order to suppress email alerts, by MVP Robert Sparnaaij: https://msunified.net/2013/11/25/lock-down-your-lync-status-and-pc-notifications-using-powershell/
        Start-Pomodoro also controls your Skype client presence, this is removed in Start-SimplePomodoro
        Get the old version here: https://github.com/janegilring/PSProductivityTools
        This function either closes Teams and starts it again or hides taskbar badges and unhides them after the session has ended for full focus on deep work
        Latest version blogged about here: https://msunified.net/2019/10/22/my-current-powershell-pomodoro-timer/
        Latest version to be found here: https://github.com/StaleHansen/Public/tree/master/Start-SimplePomodoro

        Required version: Windows PowerShell 3.0 or later 

        If you end the script prematurely, you can run the script with a 10 second lenght to reset your IFTTT and 

        It is recommended to add your Start-SimplePomodoro runline at the end of this script for easy startup

     .EXAMPLE
      Start-SimplePomodoro
     .EXAMPLE
      Start-SimplePomodoro -Minutes 15 -SpotifyPlayList spotify:playlist:XXXXXXXXXXXXXXXXXX
     .EXAMPLE
      Start-SimplePomodoro -Minutes 20 -IFTTTMuteTrigger pomodoro_start -IFTTTUnMuteTrigger pomodoro_stop -IFTTTWebhookKey XXXXXXXXX
     .EXAMPLE
      Start-SimplePomodoro -Minutes 0.1 -SpotifyPlayList spotify:playlist:XXXXXXXXXXXXXXXXXX -IFTTTMuteTrigger pomodoro_start -IFTTTUnMuteTrigger pomodoro_stop -IFTTTWebhookKey XXXXXXXXX
     .EXAMPLE
      Start-SimplePomodoro -Teamsmode Stop
     .EXAMPLE
      Start-SimplePomodoro -Secret YourFlowSecret -AutomateURI YourAutomateURI





#>

    [CmdletBinding()]
    Param (
        
        [int]$Minutes = 25, #Duration of your Pomodoro Session, default is 25 minutes
        [string]$Secret = "MySecret", #Secret for the flow trigger
        [string]$AutomateURI, #The URI used in the webrequest to your flow
        [string]$ToDoURL, #uri of your favourite spotify playlist
        [switch]$StartMusic,
        [string]$SpotifyPlayList, #uri of your favourite spotify playlist
        [string]$IFTTTMuteTrigger, #your_IFTTT_maker_mute_trigger
        [string]$IFTTTUnMuteTrigger, #your_IFTTT_maker_unmute_trigger
        [string]$IFTTTWebhookKey, #your_IFTTT_webhook_key
        [string]$StartNotificationSound = "C:\Windows\Media\Windows Proximity Connection.wav",
        [string]$EndNotificationSound = "C:\Windows\Media\Windows Proximity Notification.wav",
        [string]$Path = $env:LOCALAPPDATA+"\Microsoft\Teams\Update.exe",
        [string]$Arguments = '--processStart "Teams.exe"',
        [string]$Teamsmode = "HideBadge" #default is hide badge, set this variable to "stop" to just stop the teams client

    )

    #Clearing some space to make room for the counter

    Write-output ""
    Write-output ""
    Write-output "" 
    
    #get lenght of Pomodoro
    $minutes = Read-Host "How long is your Pomodoro?"
    Set-Clipboard -Value $minutes
    

    #Setting computer to presentation mode, will suppress most types of popups
    Write-Host "Starting presentation mode" -ForegroundColor Green
    if (Test-Path "C:\Windows\sysnative\PresentationSettings.exe"){Start-Process "C:\Windows\sysnative\PresentationSettings.exe" /start -NoNewWindow}
    else {presentationsettings /start}


    #Hide badge or stop Teams
    if ($Teamsmode -notmatch "HideBadge"){
        #Stop Microsoft Teams
        Write-Host "Closing Microsoft Teams" -ForegroundColor Green
        Get-Process -Name Teams -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue
    }
    else{
        #Hiding badges on taskbar buttons such as Outlook, Teams and ToDo
        Write-Host "Hiding badges on taskbar buttons" -ForegroundColor Green
        Set-Itemproperty -path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name 'TaskbarBadges' -value '0'
        Stop-Process -ProcessName explorer
    }


    #Start Spotify
    if ($SpotifyPlayList -ne ''){Write-Host "Opening your specified Spotify playlist" -ForegroundColor Green; Start-Process -FilePath $SpotifyPlayList}
    
    #Turn off Vibration and mute Phone using IFTTT
    if ($IFTTTMuteTrigger -ne '' -and $IFTTTWebhookKey -ne ''){
        
             try {
                      
                    $null = Invoke-RestMethod -Uri https://maker.IFTTT.com/trigger/$IFTTTMuteTrigger/with/key/$IFTTTWebhookKey -Method POST -ErrorAction Stop
                    Write-Host -Object "IFTTT mute trigger invoked successfully" -ForegroundColor Green

            }
            catch  {

                    Write-Host -Object "An error occured while invoking IFTTT mute trigger: $($_.Exception.Message)" -ForegroundColor Yellow

            }   
        
        }
    #Invoking PowerAutomate to change set current time on your Focus time calendar event, either through https trigger og manually via todo
    if ($AutomateURI -ne ''){
        $body = @()
        $body = @"
            { 
               "Duration":$Minutes,
               "Secret":"$Secret"
            }
"@
        Write-Host "Processing Focus time in your calendar and setting Teams to Focusing status" -ForegroundColor Green
        Invoke-RestMethod -Method Post -Body $Body -Uri $AutomateURI -ContentType "application/json"
    }
   elseif ($ToDoURL -ne ''){Write-Host "Opening your ToDo Pomodoro list in web, may take up to three minutes before calendar focus time update" -ForegroundColor Green; Start-Process -FilePath $ToDoURL}
   else{Write-Host "No calendar focus time specified" -ForegroundColor Green}




    #Go for deep work

    Write-Host "You are GO for flow and deep work session number $Count"
    Write-Host
    
    #Playing start sound
    if (Test-Path -Path $StartNotificationSound) {
     
        $player = New-Object System.Media.SoundPlayer $StartNotificationSound -ErrorAction SilentlyContinue
         1..2 | ForEach-Object { 
             $player.Play()
            Start-Sleep -m 3400 #invoking sleep so that the whole sound plays
        }
    }

    #Counting down to end of Pomodoro
    $seconds = $Minutes * 60
    $delay = 1 #seconds between ticks
    for ($i = $seconds; $i -gt 0; $i = $i - $delay) {
        $percentComplete = 100 - (($i / $seconds) * 100)
        Write-Progress -SecondsRemaining $i `
            -Activity "Pomodoro Focus sessions" `
            -Status "Time remaining:" `
            -PercentComplete $percentComplete
        if ($i -eq 16){Write-Host "Wrapping up, you will be available in $i seconds" -ForegroundColor Green}
        Start-Sleep -Seconds $delay
    }#Timer ended
    
    #Stopping presentation mode to re-enable outlook popups and other notifications
    Write-Host "Stopping presentation mode" -ForegroundColor Green
    if (Test-Path "C:\Windows\sysnative\PresentationSettings.exe"){Start-Process "C:\Windows\sysnative\PresentationSettings.exe" /stop -NoNewWindow}
    else {presentationsettings /stop}


    #Show badge or start Teams
    if ($Teamsmode -notmatch "HideBadge"){
        #Start Microsoft Teams again
        Write-Host "Starting Microsoft Teams" -ForegroundColor Green
        Start-Process -FilePath $Path -ArgumentList $Arguments -WindowStyle Hidden
    }
    else{
        #Show badges on taskbar buttons such as Outlook, Teams and ToDo
        Write-Host "Showing badges on taskbar buttons" -ForegroundColor Green
        Set-Itemproperty -path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name 'TaskbarBadges' -value '1'
        Stop-Process -ProcessName explorer
    }


   
    #Turn vibration on android phone back on using IFTTT
    if ($IFTTTUnMuteTrigger -ne '' -and $IFTTTWebhookKey -ne ''){

            try {
                      
                        $null = Invoke-RestMethod -Uri https://maker.IFTTT.com/trigger/$IFTTTUnMuteTrigger/with/key/$IFTTTWebhookKey -Method POST -ErrorAction Stop
           
                        Write-Host -Object "IFTTT unmute trigger invoked successfully" -ForegroundColor Green

            }
            catch  {

                Write-Host -Object "An error occured while invoking IFTTT unmute trigger: $($_.Exception.Message)" -ForegroundColor Yellow

            }   
        }

    Write-Host "Teams presence is resetting" -ForegroundColor Green
   
    #playing end notification sound
    if (Test-Path -Path $EndNotificationSound) {

    #Playing end of focus session notification
    $player = New-Object System.Media.SoundPlayer $EndNotificationSound -ErrorAction SilentlyContinue
     1..2 | ForEach-Object {
         $player.Play()
        Start-Sleep -m 1400 
    }

    }


}

$Input = "y"
$Count = 1
while ($Input -eq "y"){

#Uncomment the one of the below lines and fill in your playlist and IFTTT to have it run as part of the shortcut
#Start-SimplePomodoro -SpotifyPlayList spotify:playlist:XXXXXXXXXXXXXXXXXX -IFTTTMuteTrigger pomodoro_start -IFTTTUnMuteTrigger pomodoro_stop -IFTTTWebhookKey XXXXXXXXX -Secret YourFlowSecret -AutomateURI YourAutomateURI

Start-SimplePomodoro 

$Input = Read-Host -Prompt 'Start a new Pomodoro deep work session? (y/n)'
$Count++

}