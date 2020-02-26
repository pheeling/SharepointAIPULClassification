Param(
    [Parameter(Mandatory=$true, HelpMessage = "Please enter DataOwner as UPN")]
    [String]$dataOwner,

    [Parameter(Mandatory=$true, HelpMessage = "Please enter Customer Tenant ID")]
    [String]$tenantID,

    [Parameter(Mandatory=$true, HelpMessage = "Which LabelID should be used for classification?")]
    [String]$labelId,

    [Parameter(Mandatory=$true, HelpMessage = "Define Path to Sharepoint Libraries List Text File?")]
    [String]$documentLibraryList,

    [Parameter(Mandatory=$true, HelpMessage = "Define WebApp ID for connection?")]
    [String]$webAppID,
    
    [Parameter(Mandatory=$true, HelpMessage = "Sharepoint URL?")]
    [uri]$sharepointLoginUrl
)

$Global:srcPath = split-path -path $MyInvocation.MyCommand.Definition 
$Global:mainPath = split-path -path $srcPath
$Global:resourcespath = join-path -path "$mainPath" -ChildPath "resources"
$Global:errorVariable = "Stop"
$Global:logFile = "$resourcespath\processing.log"
$webAppKeyXMLFile = "..\ressources\${env:USERNAME}_appKey_$($tenantID).xml"

Import-Module -Force "$resourcespath\ErrorHandling.psm1"
Import-Module -Force "$resourcespath\SharepointClassification.psm1"

"$(Get-Date) [Processing] Start--------------------------" >> $Global:logFile

$sharepoint = Get-NewSharepointClassification($tenantID)

#Requirements Check
try {
    if ((Get-InstalledModule -name  Microsoft.Online.Sharepoint.Powershell | 
        Select-Object -ExpandProperty version) -ge 16.0.19724.12000) {
        "$(Get-Date) [RequirementsCheck] Module Sharepoint exists" >> $Global:logFile
    }
    else {
        Install-Module -name Microsoft.Online.Sharepoint.Powershell -AllowClobber -Force
    }
    if(Get-Command Get-AIPFileStatus -ErrorAction SilentlyContinue){
        "$(Get-Date) [RequirementsCheck] Module AIP exists" >> $Global:logFile
    }
    if(Test-Path $webAppKeyXMLFile){
        "$(Get-Date) [RequirementsCheck] token exists" >> $Global:logFile
        $webAppKey = [PSCredential] (Import-Clixml $webAppKeyXMLFile).GetNetworkCredential().Password
    } else {
        Get-Credential | Export-Clixml -Path $webAppKeyXMLFile
        $webAppKey = [PSCredential] (Import-Clixml $webAppKeyXMLFile).GetNetworkCredential().Password
        $sharepoint.connectAIPService($webAppID,$webAppKey)
    }
} catch {
    "$(Get-Date) [RequirementsCheck] Module installation failed: $PSItem" >> $Global:logFile
    #Get-NewErrorHandling "$(Get-Date) [RequirementsCheck] Module installation failed" $PSItem
}

$sharepoint.connectSPO($sharepointLoginUrl)
$arrayLibrary = $sharepoint.readDocumentLibraryListFile($documentLibraryList)
$sharepoint.getSharepointLibraryEntries($arrayLibrary)
$sharepoint.fileClassification($sharepoint.documentsList, $labelId, $dataOwner, $sharepointLoginUrl)

"$(Get-Date) [Processing] Stopped -----------------------" >> $Global:logFile

$sharepoint.fileRetention($Global:logfile)
$sharepoint.fileRetention("$Global:resourcespath\AIPResultStatus.csv")
$sharepoint.fileRetention("$Global:resourcespath\AIPFilestatus.csv")