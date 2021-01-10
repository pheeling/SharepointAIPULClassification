Param(
    [Parameter(Mandatory=$true, HelpMessage = "Please enter DataOwner as UPN")]
    [String]$dataOwner,

    [Parameter(Mandatory=$true, HelpMessage = "Please enter Customer Tenant ID")]
    [String]$tenantID,

    [Parameter(Mandatory=$true, HelpMessage = "Which LabelID should be used for classification?")]
    [String]$labelId,

    [Parameter(Mandatory=$true, HelpMessage = "Define Path to Sharepoint Libraries List Text File?")]
    [String]$documentLibraryList,

    [Parameter(Mandatory=$true, HelpMessage = "Define Path to Sharepoint Libraries List Text File?")]
    [String]$filterFileList,

    [Parameter(Mandatory=$false, HelpMessage = "Define WebApp ID for connection?")]
    [String]$webAppID,
    
    [Parameter(Mandatory=$true, HelpMessage = "Sharepoint URL?")]
    [uri]$sharepointLoginUrl
)

$Global:srcPath = split-path -path $MyInvocation.MyCommand.Definition 
$Global:mainPath = split-path -path $srcPath
$Global:resourcespath = join-path -path "$mainPath" -ChildPath "resources"
$Global:errorVariable = "Stop"
$Global:logFile = "$resourcespath\processing.log"
#$webAppKeyXMLFile = "$Global:resourcespath\${env:USERNAME}_appKey_$($tenantID).xml"

Import-Module -Force "$resourcespath\ErrorHandling.psm1"
Import-Module -Force "$resourcespath\SharepointClassification.psm1"

"$(Get-Date) [Processing] Start--------------------------" >> $Global:logFile

$sharepoint = Get-NewSharepointClassification($tenantID)

#Requirements Check
try {
    if (Get-Command Connect-PnPOnline) {
        "$(Get-Date) [RequirementsCheck] Module Sharepoint exists" >> $Global:logFile
    } else {
        Install-Module SharePointPnPPowerShellOnline
    }
    if(Get-Command Connect-AipService -ErrorAction SilentlyContinue){
        "$(Get-Date) [RequirementsCheck] Module AIP Classic exists" >> $Global:logFile
    } else {
        #Dependent on Docker Image
        Import-Module "C:\source\AIP.dll"
    }
    if(Get-Command Connect-IPPSSession -ErrorAction SilentlyContinue){
        "$(Get-Date) [RequirementsCheck] Module ExchangeOnlineManagement exists" >> $Global:logFile
    } else {
        #Dependent on Docker Image
        Install-Module ExchangeOnlineManagement
    }
    if(Get-Command Get-AIPFileStatus -ErrorAction SilentlyContinue){
        "$(Get-Date) [RequirementsCheck] Module AIP exists" >> $Global:logFile
    } else {
        #Dependent on Docker Image
        Import-Module "C:\source\AIP.dll"
    }

    
} catch {
    "$(Get-Date) [RequirementsCheck] Module installation failed: $PSItem" >> $Global:logFile
    #Get-NewErrorHandling "$(Get-Date) [RequirementsCheck] Module installation failed" $PSItem
}

$sharepoint.connectAIPServiceClassic()
$sharepoint.connectAIPServiceUL()
$sharepoint.connectSPO($sharepointLoginUrl)
$arrayLibrary = $sharepoint.readTextInputs($documentLibraryList)
$arrayFilter = $sharepoint.readTextInputs($filterFileList)
$sharepoint.getSharepointLibraryEntries($arrayLibrary)
$sharepoint.fileClassification($sharepoint.documentsList, $labelId, $dataOwner, $sharepointLoginUrl, $arrayFilter)

"$(Get-Date) [Processing] Stopped -----------------------" >> $Global:logFile

$sharepoint.fileRetention($Global:logfile)
$sharepoint.fileRetention("$Global:resourcespath\AIPResultStatus.csv")
$sharepoint.fileRetention("$Global:resourcespath\AIPFilestatus.csv")