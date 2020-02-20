function Get-NewSharepointClassification($customerAbbreviations){
    return [Authentication]::new()
}

class SharepointClassification {

    [String] $credentialFilepath = "..\ressources\${env:USERNAME}_cred_$customerAbbreviations.xml"
    [PSCredential] $credentials

    SharepointClassification(){
        if (!(Test-Path $this.credentialFilepath)) {
            Get-Credential | Export-Clixml -Path $this.credentialFilepath
        }
        $this.credentials = Import-Clixml $this.credentialFilepath
    }

    connectSPO([uri] $sharepointLoginURL){
        Connect-PnPOnline -url  $sharepointLoginUrl -credential $this.credentials
    }

    connectAIPService([String] $webAppID, [String] $webAppKey, [String] $nativeAppID){
        $token = (Set-AIPAuthentication -webAppId $webAppID -webAppKey $webAppKey -nativeAppId $nativeAppID).token
        Set-AIPAuthentication -webAppId $webAppID -webAppKey $webAppKey -nativeAppId $nativeAppID -token $token
    }

    fileClassification($documentLibraryList, $labelId, $dataOwner, $sharepointLoginUrl){
        foreach ($item in $documentLibraryList) { 
            if (($item.Fieldvalues.FileRef -like "*.docx") -or ($item.Fieldvalues.FileRef -like "*.xlsx") -or ($item.Fieldvalues.FileRef -like "*.pptx")) { 
                Get-PnPFile -Url $item.FieldValues.FileRef -Path $Global:resourcespath -AsFile -Force
                $element = "$($Global:resourcespath)\$($item.Fieldvalues.FileLeafRef)"
        
                Write-host $element
                $element | Set-AIPFileLabel -LabelId $labelId -Owner $dataOwner -PreserveFileDetails | Export-Csv -Append "$Global:resourcespath\AIPResultStatus.csv"
                $element | Get-AIPFileStatus | Export-csv -Append "$Global:resourcespath\AIPFilestatus.csv"
                Add-PnPFile -Path "$($element)" -Folder $item.FieldValues.FileDirRef.Substring($sharepointLoginUrl.AbsolutePath.Length) -Checkout
                Remove-Item -Path "$($element)" -Confirm:$false -Force    
            }
        }
    }
}