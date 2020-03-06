function Get-NewSharepointClassification($tenantID){
    return [SharepointClassification]::new($tenantID)
}

class SharepointClassification {

    [String] $tenantID
    [String] $sharepointCredentialFilepath = "$Global:resourcespath\${env:USERNAME}_sharepointCred_$($tenantID).xml"
    [PSCredential] $sharepointCredentials
    [String] $svcCredentialFilepath = "$Global:resourcespath\${env:USERNAME}_svcCred_$($tenantID).xml"
    [PSCredential] $svcCredentials
    [Array] $documentsList = @()

    SharepointClassification($tenantID){
        if (!(Test-Path $this.sharepointCredentialFilepath)) {
            Get-Credential | Export-Clixml -Path $this.sharepointCredentialFilepath
        }
        $this.sharepointCredentials = Import-Clixml $this.sharepointCredentialFilepath
        $this.tenantID = $tenantID
    }

    connectSPO([uri] $sharepointLoginURL){
        #Connect-PnPOnline -url  $sharepointLoginUrl -credential $this.sharepointCredentials
        Connect-PnPOnline -url $sharepointLoginUrl -UseWebLogin
    }

    connectAIPService([String] $webAppID, [String] $webAppKey){
        if (!(Test-Path $this.svcCredentialFilepath)) {
            Get-Credential | Export-Clixml -Path $this.svcCredentialFilepath
            $this.svcCredentials = Import-Clixml $this.svcCredentialFilepath
            Write-Error "If your running this script the first time, please run Set-AIPAuthentication with elevated Privileges, 
            Set-AIPAuthentication -AppId $webAppID -AppSecret $webAppKey -TenantId $this.tenantID -DelegatedUser $this.sharepointCredentials.UserName -OnBehalfOf $this.svcCredentials"
            Exit 1
        }
        
    }

    [string[]] readDocumentLibraryListFile($documentLibraryList){
        return Get-Content -Path $documentLibraryList
    }

    getSharepointLibraryEntries($array){
        foreach ($entry in $array){
            $this.documentsList += Get-PnPListItem -List $entry -PageSize 100
        }
    }

    fileClassification($documentsList, $labelId, $dataOwner, $sharepointLoginUrl){
        foreach ($item in $documentsList) {
            if (($item.Fieldvalues.FileRef -like "*.docx") -or ($item.Fieldvalues.FileRef -like "*.xlsx") -or ($item.Fieldvalues.FileRef -like "*.pptx")) { 
                Get-PnPFile -Url $item.FieldValues.FileRef -Path $Global:resourcespath -AsFile -Force
                $element = "$($Global:resourcespath)\$($item.Fieldvalues.FileLeafRef)"
                $modifiedBy = Get-PnPUser | Where-Object {$_.Title -eq $item.FieldValues.Editor.LookupValue}
                $modifiedDate = $($item.FieldValues.Modified)
        
                Write-host $element
                "$(Get-Date) [SharepointClassification] File for Classification $element : $($item.Fieldvalues.FileRef)" >> $Global:logFile
                $element | Get-AIPFileStatus | Where-Object {$_.IsLabeled -eq $false} | Set-AIPFileLabel -LabelId $labelId -Owner $dataOwner -PreserveFileDetails | Export-Csv -Append "$Global:resourcespath\AIPResultStatus.csv"
                $element | Get-AIPFileStatus | Export-csv -Append "$Global:resourcespath\AIPFilestatus.csv"
                Add-PnPFile -Path "$($element)" -Folder $item.FieldValues.FileDirRef.Substring($sharepointLoginUrl.AbsolutePath.Length) -Values @{Editor=$modifiedBy.Email;Modified=$modifiedDate} -Checkout
                Remove-Item -Path "$($element)" -Confirm:$false -Force
            }
        }
    }

    fileRetention($filepath){
        if ((Get-ChildItem -path $filepath).Length -gt 5242880) {
            Remove-Item -Path $filepath
        }
    }
}