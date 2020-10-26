<#
    .DESCRIPTION
        An runbook that;
        authenticates with a certificate
        Gets its input parameter from Logic Apps 
        Iterates through an object from the input then stores each item in an ArrayList
        disables external sharing configuration for each onedrive owner in the list

    .NOTES
        AUTHOR: Abioye Mohammed
        LASTEDIT: Oct. 06, 2020
#>

#Create an oject parameter to receive input from Power Automate
param
    (
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] 
            [Object] 
            $UnResponsiveUsers = "All"
    )

#create an empty arraylist to store the emails of users whose onedrive sites will be disabled
$ArrayOfUsers = [System.Collections.ArrayList]::new()

#iterate through the input parameter email object then store each item in the Arraylist
foreach($person in $UnResponsiveUsers.value.email) 
{ 
    $ArrayOfUsers.add("$person")
} 

#get Automation account name
$connectionName = "AzureRunAsConnection"

# Get the connection "AzureRunAsConnection "
$servicePrincipalConnection = Get-AutomationConnection -Name $connectionName 

#logOn credentials
    $tenant               = "Enter your Domain Name"                               # O365 TENANT NAME
    $clientId             = $servicePrincipalConnection.ApplicationID   # AAD APP PRINCIPAL CLIENT ID

#Stored as a variable
    $appPrincipalPwdVar   = 'Enter the name of the password variable stored in Automation account or enter a string'  # CERT PASSWORD VARIABLE

#stored as a certificate
    $appPrincipalCertVar  = 'Enter the name of the .pfx certificate uploaded to your automation account > certicates tab'                          # CERT NAME VARIABLE

$VerbosePreference = "Continue"

# load the saved automation properties
    $appPrincipalCertificatePwd = Get-AutomationVariable    -Name $appPrincipalPwdVar
    $appPrincipalCertificate    = Get-AutomationCertificate -Name $appPrincipalCertVar
    
# load the cert from automation store and save it locally so it can be used by the PnP cmdlets
    # temp path to store cert
    $certificatePath = "C:\temp-certificate-$([System.Guid]::NewGuid().ToString()).pfx" 
    $appPrincipalCertificateSecurePwd = ConvertTo-SecureString -String $appPrincipalCertificatePwd -AsPlainText -Force
    Export-PfxCertificate -FilePath $certificatePath -Password $appPrincipalCertificateSecurePwd -Cert $appPrincipalCertificate

# connect to the tenant admin site
    Write-Verbose -Message "$(Get-Date) - Connecting to https://$tenant-admin.sharepoint.com"
    Connect-PnPOnline `
                -Url                 "https://$tenant-admin.sharepoint.com" `
                -Tenant              "$tenant.onmicrosoft.com" `
                -ClientId            $clientId `
                -CertificatePath     $certificatePath `
                -CertificatePassword $appPrincipalCertificateSecurePwd

# delete the local cert
    Write-Verbose -Message "$(Get-Date) - Deleting Certificate"
    Remove-Item -Path $certificatePath -Force -ErrorAction SilentlyContinue
        
# get the onedrive site url for each user
    $token = Get-PnPGraphAccessToken
    $mappings = @()
    foreach( $user in $ArrayOfUsers )
    {
        $json = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$user/drive" -Headers @{ Authorization = "Bearer $token"; Accept = "application/json" }
        
        $mappings += [PSCustomObject] @{
            UserName = $user
            DriveUrl = $json.webUrl -replace "/Documents$", ""
        }
    }

# set external sharing configuration for each onedrive owner
    foreach( $mapping in $mappings )
    {
        if( -not [string]::IsNullOrWhiteSpace($mapping.DriveUrl) )
        {
            Set-PnPTenantSite -Url $mapping.DriveUrl -SharingCapability Disabled
        }
        $mapping.DriveUrl
    }
    #completed execution
