function Get-NewSharepointClassification($customerAbbreviations){
    return [SharepointClassification]::new()
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