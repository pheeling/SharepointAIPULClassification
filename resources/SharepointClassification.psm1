function Get-NewSharepointClassification($tenantID){
    return [SharepointClassification]::new($tenantID)
}

class SharepointClassification {

    [String] $tenantID
    [String] $credentialFilepath = "..\ressources\${env:USERNAME}_cred_$($this.tenantID).xml"
    [PSCredential] $credentials

    SharepointClassification($tenantID){
        if (!(Test-Path $this.credentialFilepath)) {
            Get-Credential | Export-Clixml -Path $this.credentialFilepath
        }
        $this.credentials = Import-Clixml $this.credentialFilepath
        $this.tenantID = $tenantID
    }

    connectSPO([uri] $sharepointLoginURL){
        Connect-PnPOnline -url  $sharepointLoginUrl -credential $this.credentials
    }

    connectAIPService([String] $webAppID, [String] $webAppKey, [String] $nativeAppID){
        Set-AIPAuthentication -AppId $webAppID -AppSecret $webAppKey -TenantId $this.tenantID -DelegatedUser $this.credentials.UserName -OnBehalfOf $this.credentials
    }

    [string[]] readDocumentLibraryListFile($documentLibraryList){
        return Get-Content -Path $documentLibraryList
    }

    fileClassification($arrayLibrary, $labelId, $dataOwner, $sharepointLoginUrl){
        foreach ($item in $arrayLibrary) { 
            if (($item.Fieldvalues.FileRef -like "*.docx") -or ($item.Fieldvalues.FileRef -like "*.xlsx") -or ($item.Fieldvalues.FileRef -like "*.pptx")) { 
                Get-PnPFile -Url $item.FieldValues.FileRef -Path $Global:resourcespath -AsFile -Force
                $element = "$($Global:resourcespath)\$($item.Fieldvalues.FileLeafRef)"
        
                Write-host $element
                "$(Get-Date) [SharepointClassification] File for Classification $element : $($item.Fieldvalues.FileRef)" >> $Global:logFile
                $element | Set-AIPFileLabel -LabelId $labelId -Owner $dataOwner -PreserveFileDetails | Export-Csv -Append "$Global:resourcespath\AIPResultStatus.csv"
                $element | Get-AIPFileStatus | Export-csv -Append "$Global:resourcespath\AIPFilestatus.csv"
                Add-PnPFile -Path "$($element)" -Folder $item.FieldValues.FileDirRef.Substring($sharepointLoginUrl.AbsolutePath.Length) -Checkout
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