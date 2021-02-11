break
#By Ståle Hansen, Twitter: @StaleHansen
#runs as code and not as a script

###############################################################################
## Script to merge compliance score and secure score and adding custom attributes
###############################################################################


$Folder = "C:\Temp\"
$ComplianceScore = "Compliance Manager - Microsoft 365 compliance.csv"
$SecureScore = "Microsoft Secure Score - Microsoft 365 security.csv"

$NewComplianceScore = "New Compliance Manager - Microsoft 365 compliance.csv"
$NewSecureScore = "New Microsoft Secure Score - Microsoft 365 security.csv"




###############################################################################
## Create an object that has all the attributes we want
###############################################################################

function Add-ExtendedAttributeList{

    New-Object -TypeName PSCustomObject -Property @{
        ImprovementAction = $null
        Priority = $null
        Effort = $null
        Responsible = $null
        TechnicalResource = $null
        Comment = $null
        ImplementationStatus = $null
        TestStatus = $null
        TestedDate = $Null
        ScoreUpdated = $null
        Source = $null
        Solutions = $null
        Categories = $null
        PointsAchieved = $Null
        ActionType = $Null
        Regulations = $null
        Group = $null
        Assessments = $null
        DateImported = $null
        LastSynced = $null
        MicrosoftUpdate = $null
        ScoreImpact = $null
        Regressed = $null
        HaveLicense = $null
        Description = $null
        Documentation = $null
    }

}

###############################################################################
## Immport the scores and create the complete score list
###############################################################################

#Importing the score csv's
$ComplianceScoreImport = Import-Csv -Path $Folder$ComplianceScore
$SecureScoreImport = Import-Csv -Path $Folder$SecureScore


$CompleteScore = @()


#creating the extended compliance score list
    foreach ($Line in $ComplianceScoreImport){
                $addObject = Add-ExtendedAttributeList
                $addObject.ImprovementAction = $Line.'Improvement action'
                $addObject.Source = "Compliance Score"
                $addObject.Solutions = $Line.Solutions
                $addObject.Categories = $Line.Categories
                $addObject.PointsAchieved = $Line.'Points achieved'
                $addObject.ActionType = $Line.'Action Type'
                $addObject.Group = $Line.Group
                $addObject.Assessments = $Line.Assessments
                $addObject.DateImported = (Get-Date -Format "yyyy-MM-dd")
                $CompleteScore += $addObject
            
        }


#creating the extended secure score list
    foreach ($Line in $SecureScoreImport){
                $addObject = Add-ExtendedAttributeList
                $addObject.ImprovementAction = $Line.'Improvement action'
                $addObject.Comment = $Line.Notes
                $addObject.Source = "Secure Score"
                $addObject.Solutions = $Line.Product
                $addObject.Categories = $Line.Categories
                $addObject.PointsAchieved = $Line.'Points achieved'
                $addObject.ActionType = $Line.'Action Type'
                $addObject.Group = $Line.Group
                $addObject.Assessments = $Line.Assessments
                $addObject.DateImported = (Get-Date -Format "yyyy-MM-dd")
                $addObject.LastSynced = $Line.'Last Synced'
                $addObject.MicrosoftUpdate = $Line.'Microsoft update'
                $addObject.ScoreImpact = $Line.'Score impact'
                $addObject.Regressed = $Line.Regressed
                $addObject.HaveLicense = $Line.'Have License?'
                $addObject.Description = $Line.Description
                $addObject.ActionType = "Technical"
                $CompleteScore += $addObject
            
        }

$CompleteScore.ImprovementAction
$CompleteScore.Count

$CompleteScore | Select-Object -Property ImprovementAction, Priority, Effort, Responsible, TechnicalResource, `
Comment, ImplementationStatus, TestStatus, TestedDate, ScoreUpdated, Source, Solutions, Categories, ActionType, `
Regulations, Group, Assessments, DateImported, PointsAchieved, LastSynced, MicrosoftUpdate, ScoreImpact, Regressed, `
HaveLicense, Documentation, Description | Export-Csv -Path $Folder"CompleteScore.csv" -NoTypeInformation -Encoding UTF8


###############################################################################
## Find new score actions
###############################################################################

$Folder = "C:\Temp\"


#Importing the old compliance score csv's
$ComplianceScoreImport = Import-Csv -Path $Folder$ComplianceScore

#Importing the new compliance score csv's
$NewComplianceScoreImport = Import-Csv -Path $Folder$NewComplianceScore

$FinalAddedActions = @()
$addedCompleteScore = @()

if ($ComplianceScoreImport.count -lt $NewComplianceScoreImport.count){

    $compared = Compare-Object -ReferenceObject $ComplianceScoreImport.'Improvement action' -DifferenceObject $NewComplianceScoreImport.'Improvement action' | Where-Object {$_.SideIndicator -match "=>"}
    $count = $compared.count; $count--
    $FinalAddedActions = @()
    while ($count -ge 0){
        $AddedActions = $NewComplianceScoreImport | Where-Object {$_.'Improvement action' -match $compared.InputObject[$Count]}
        $count--
        $FinalAddedActions += $AddedActions 
    }

    Write-host Number of new actions: @($FinalAddedActions).count

    #creating the extended compliance score list
    foreach ($Line in $FinalAddedActions){
                $addObject = Add-ExtendedAttributeList
                $addObject.ImprovementAction = $Line.'Improvement action'
                $addObject.Source = "Compliance Score"
                $addObject.Solutions = $Line.Solutions
                $addObject.Categories = $Line.Categories
                $addObject.PointsAchieved = $Line.'Points achieved'
                $addObject.ActionType = $Line.'Action Type'
                $addObject.Group = $Line.Group
                $addObject.Assessments = $Line.Assessments
                $addObject.DateImported = (Get-Date -Format "yyyy-MM-dd")
                $addedCompleteScore += $addObject
            
        }
        $addedCompleteScore.count
}
else{Write-host No new actions detected}


#Importing the old secure score csv's
$SecureScoreImport = Import-Csv -Path $Folder$SecureScore

#Importing the new secure score csv's
$NewSecureScoreImport = Import-Csv -Path $Folder$NewSecureScore


if ($SecureScoreImport.count -lt $NewSecureScoreImport.count){

    $compared = Compare-Object -ReferenceObject $SecureScoreImport.'Improvement action' -DifferenceObject $NewSecureScoreImport.'Improvement action' | Where-Object {$_.SideIndicator -match "=>"}
    $count = $compared.count; $count--
    if ($FinalAddedActions -eq $Null){$FinalAddedActions = @()}
    while ($count -ge 0){
        $AddedActions = $NewSecureScoreImport | Where-Object {$_.'Improvement action' -match $compared.InputObject[$Count]}
        $count--
        $FinalAddedActions += $AddedActions 
    }

    Write-host Number of new actions: @($FinalAddedActions).count

    #creating the extended secure score list
       foreach ($Line in $FinalAddedActions){
                $addObject = Add-ExtendedAttributeList
                $addObject.ImprovementAction = $Line.'Improvement action'
                $addObject.Comment = $Line.Notes
                $addObject.Source = "Secure Score"
                $addObject.Solutions = $Line.Product
                $addObject.Categories = $Line.Categories
                $addObject.PointsAchieved = $Line.'Points achieved'
                $addObject.ActionType = $Line.'Action Type'
                $addObject.Group = $Line.Group
                $addObject.Assessments = $Line.Assessments
                $addObject.DateImported = (Get-Date -Format "yyyy-MM-dd")
                $addObject.LastSynced = $Line.'Last Synced'
                $addObject.MicrosoftUpdate = $Line.'Microsoft update'
                $addObject.ScoreImpact = $Line.'Score impact'
                $addObject.Regressed = $Line.Regressed
                $addObject.HaveLicense = $Line.'Have License?'
                $addObject.Description = $Line.Description
                $addObject.ActionType = "Technical"
                $AddedCompleteScore += $addObject
            
        }
        $AddedCompleteScore.count

}
else{Write-host No new actions detected}

$addedCompleteScore.ImprovementAction
$addedCompleteScore.Count

$addedCompleteScore | Select-Object -Property ImprovementAction, Priority, Effort, Responsible, TechnicalResource, `
Comment, ImplementationStatus, TestStatus, TestedDate, ScoreUpdated, Source, Solutions, Categories, ActionType, `
Regulations, Group, Assessments, DateImported, PointsAchieved, LastSynced, MicrosoftUpdate, ScoreImpact, Regressed, `
HaveLicense, Documentation, Description | Export-Csv -Path $Folder"AddedCompleteScore.csv" -NoTypeInformation -Encoding UTF8
