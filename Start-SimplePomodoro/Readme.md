# Start-SimplePomodoro

Start-SimplePomodoro is a function, add your variables and the run syntax in the ps1 file for easy reusability

## Script Result

![Start-SimplePomodoro](https://github.com/StaleHansen/Public/blob/master/Start-SimplePomodoro/Start-SimplePomodoro.png)

## SYNOPSIS
      Start-SimplePomodoro is a function command to start a new Pomodoro session with additional actions. This is a simplified version of the Start-Pomodoro
      Read the latest updates about this script: https://msunified.net/?s=pomodoro
## DESCRIPTION
        By MVP Ståle Hansen (http://msunified.net) with modifications by Jan Egil Ring (https://github.com/janegilring)
        Pomodoro function by Nathan.Run() http://nathanhoneycutt.net/blog/a-pomodoro-timer-in-powershell/
        Note: for desktops you need to enable presentation settings in order to suppress email alerts, by MVP Robert Sparnaaij: https://msunified.net/2013/11/25/lock-down-your-lync-status-and-pc-notifications-using-powershell/
        Start-Pomodoro also controls your Skype client presence, this is removed in Start-SimplePomodoro
        This function either closes Teams and starts it again or hides taskbar badges and unhides them after the session has ended for full focus on deep work
        Latest version blogged about here: https://msunified.net/2019/10/22/my-current-powershell-pomodoro-timer/
        Latest version to be found here: https://github.com/StaleHansen/Public/tree/master/Start-SimplePomodoro
        Create shortcut, review this blogpost: https://msunified.net/2020/06/14/updated-my-pomodoro-powershell-timer/ 
        Required version: Windows PowerShell 3.0 or later 
        If you end the script prematurely, you can run the script with a 10 second lenght to reset your IFTTT and 
        It is recommended to add your Start-SimplePomodoro runline at the end of this script for easy startup
  ## Examples      
     .EXAMPLE
      Start-SimplePomodoro
     .EXAMPLE
      Start-SimplePomodoro -Minutes 15 -SpotifyPlayList spotify:playlist:XXXXXXXXXXXXXXXXXX
     .EXAMPLE
      Start-SimplePomodoro -Minutes 20 -IFTTMuteTrigger pomodoro_start -IFTTUnMuteTrigger pomodoro_stop -IFTTWebhookKey XXXXXXXXX
      .EXAMPLE
      Start-SimplePomodoro -Minutes 0.1 -SpotifyPlayList spotify:playlist:XXXXXXXXXXXXXXXXXX -IFTTMuteTrigger pomodoro_start -IFTTUnMuteTrigger pomodoro_stop -IFTTWebhookKey XXXXXXXXX
      .EXAMPLE
      Start-SimplePomodoro -Teamsmode Stop

