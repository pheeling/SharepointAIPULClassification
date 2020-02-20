Param(
    [Parameter(Mandatory=$true, HelpMessage = "Please enter DataOwner as UPN")]
    [String]$dataOwner,

    [Parameter(Mandatory=$true, HelpMessage = "Please enter Customer Abbreviation e.g DT aka Dinotronic")]
    [String]$customerAbbreviations,

    [Parameter(Mandatory=$true, HelpMessage = "Which LabelID should be used for classification?")]
    [String]$labelId,

    [Parameter(Mandatory=$true, HelpMessage = "Define Path to Sharepoint Libraries List File?")]
    [String]$documentLibraryList,

    [Parameter(Mandatory=$true, HelpMessage = "Sharepoint URL?")]
    [uri]$sharepointLoginUrl
)

$Global:srcPath = split-path -path $MyInvocation.MyCommand.Definition 
$Global:mainPath = split-path -path $srcPath
$Global:resourcespath = join-path -path "$mainPath" -ChildPath "resources"
$Global:errorVariable = "Stop"
$Global:logFile = "$resourcespath\processing.log"

Import-Module -Force "$resourcespath\ErrorHandling.psm1"

"$(Get-Date) [Processing] Start--------------------------" >> $Global:logFile
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
} catch {
    "$(Get-Date) [RequirementsCheck] Module installation failed: $PSItem" >> $Global:logFile
    #Get-NewErrorHandling "$(Get-Date) [RequirementsCheck] Module installation failed" $PSItem
}

"$(Get-Date) [Processing] Stopped -----------------------" >> $Global:logFile
if ((Get-ChildItem -path $Global:logfile).Length -gt 5242880) {
    Remove-Item -Path $Global:logFile
}