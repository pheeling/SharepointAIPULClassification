function Get-NewSharepointClassification($tenantID){
    return [SharepointClassification]::new($tenantID)
}

class SharepointClassification {

    [String] $tenantID
    [String] $sharepointCredentialFilepath = "$Global:resourcespath\${env:USERNAME}_sharepointCred_$($tenantID).xml"
    [PSCredential] $sharepointCredentials
    [String] $svcCredentialFilepath = "$Global:resourcespath\${env:USERNAME}_svcCred_$($tenantID).xml"
    [String] $AIPPKMCredentialFilepath = "$Global:resourcespath\${env:USERNAME}_AIP_PKM_Cred_$($tenantID).xml"
    [String] $AIPEXACredentialFilepath = "$Global:resourcespath\${env:USERNAME}_AIP_EXA_Cred_$($tenantID).xml"
    [String] $AIPDelegatedUserFilePath = "$Global:resourcespath\${env:USERNAME}_AIPDelegatedUser_Cred_$($tenantID).xml"
    [PSCredential] $svcCredentials
    [PSCredential] $AIPPKMCredentials
    [PSCredential] $AIPEXACredentials
    [PScredential] $AIPDelegatedUserCredentials
    [Array] $documentsList = @()

    SharepointClassification($tenantID){
        if (!(Test-Path $this.sharepointCredentialFilepath)) {
            $this.createCredentials($this.sharepointCredentialFilepath) | Export-Clixml -Path $this.sharepointCredentialFilepath
        }
        if (!(Test-Path $this.AIPPKMCredentialFilepath)) {
            $this.createCredentials($this.AIPPKMCredentialFilepath) | Export-Clixml -Path $this.AIPPKMCredentialFilepath
        }
        if (!(Test-Path $this.AIPEXACredentialFilepath)) {
            $this.createCredentials($this.AIPEXACredentialFilepath) | Export-Clixml -Path $this.AIPEXACredentialFilepath
        }
        if (!(Test-Path $this.svcCredentialFilepath)) {
            $this.createCredentials($this.AIPDelegatedUserFilePath) | Export-Clixml -Path $this.AIPDelegatedUserFilePath
            $this.createCredentials($this.svcCredentialFilepath) | Export-Clixml -Path $this.svcCredentialFilepath
            $this.svcCredentials = Import-Clixml $this.svcCredentialFilepath
        }
        $this.sharepointCredentials = Import-Clixml $this.sharepointCredentialFilepath
        $this.AIPPKMCredentials = Import-Clixml $this.AIPPKMCredentialFilepath
        $this.AIPEXACredentials = Import-Clixml $this.AIPEXACredentialFilepath
        $this.tenantID = $tenantID
        Set-AIPAuthentication -AppId $this.AIPDelegatedUserCredentials.UserName -AppSecret `
        $this.AIPDelegatedUserCredentials.GetNetworkCredential().Password `
        -TenantId $this.tenantID -DelegatedUser $this.sharepointCredentials.UserName -OnBehalfOf $this.svcCredentials
    }

    [PScredential] createCredentials($username,$pass){
        $secpassword = ConvertTo-SecureString $pass -AsPlainText -Force
        return New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $secpassword)
    }

    [PScredential] createCredentials($xmlfilename){
        $username = Read-Host -Prompt "Input your username for XML$($xmlfilename)::"
        $pass = Read-Host -Prompt 'Input your password::'
        If($pass){
            $secpassword = ConvertTo-SecureString $pass -AsPlainText -Force
            return New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $secpassword)
        } else {
            return New-Object System.Management.Automation.PSCredential ($username, (new-object System.Security.SecureString))
        }
    }

    connectSPO([uri] $sharepointLoginURL){
        #Connect-PnPOnline -url  $sharepointLoginUrl -UseWebLogin
        Connect-PnPOnline -url $sharepointLoginUrl -Credentials $this.sharepointCredentials
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

    connectAIPServiceClassic(){
        Connect-AipService -Credential $this.AIPPKMCredentials
    }

    connectAIPServiceUL(){
        Connect-IPPSSession -Credential $this.AIPEXACredentials
    }

    [string[]] readTextInputs($documentLibraryList){
        return Get-Content -Path $documentLibraryList
    }

    getSharepointLibraryEntries($array){
        foreach ($entry in $array){
            $this.documentsList += Get-PnPListItem -List $entry -PageSize 100
        }
    }

    fileClassification($documentsList, $labelId, $dataOwner, $sharepointLoginUrl, $filter){
        foreach ($item in $documentsList) {
            if ($filter -contains $item.Fieldvalues.FileLeafRef) { 
                Get-PnPFile -Url $item.FieldValues.FileRef -Path $Global:resourcespath -AsFile -Force
                $element = "$($Global:resourcespath)\$($item.Fieldvalues.FileLeafRef)"
                $modifiedBy = Get-PnPUser | Where-Object {$_.Title -eq $item.FieldValues.Editor.LookupValue}
                $modifiedDate = $($item.FieldValues.Modified)
        
                Write-host $element
                "$(Get-Date) [SharepointClassification] File for Classification $element : $($item.Fieldvalues.FileRef)" >> $Global:logFile
                $element | Set-AIPFileLabel -RemoveLabel
                $element | Set-AIPFileLabel -RemoveProtection
                Set-AIPFileLabel -LabelId $labelId -Owner $dataOwner -PreserveFileDetails | Export-Csv -Append "$Global:resourcespath\AIPResultStatus.csv"
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