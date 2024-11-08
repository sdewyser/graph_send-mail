# Graph_Send-Mail.psm1

# Define the module manifest
function Get-ModuleManifest {
    @{
        RootModule = 'Graph_Send-Mail.psm1'
        ModuleVersion = '1.1.0'
        GUID = '1e4d7d4d-3a8d-4d5e-9f6d-e82ffde1cbb2'
        Author = 'Stefaan Dewulf'
        Description = 'A PowerShell module to send emails with attachments using Microsoft Graph API'
        CompanyName = 'dewyser.net'
    }
}

# Import necessary modules
Import-Module -Name Microsoft.PowerShell.SecretManagement -ErrorAction SilentlyContinue

# Define variables for Microsoft Graph API
$global:GraphAPIUrl = "https://graph.microsoft.com/v1.0"

# Function to get an access token using client credentials
function Get-AccessToken {
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory)]
        [string]$ClientSecret
    )

    $body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    $response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
    return $response.access_token
}

# Function to send an email with optional attachments via Microsoft Graph API
function Send-Mail {
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,
        [Parameter(Mandatory)]
        [string]$ClientId,
        [Parameter(Mandatory)]
        [string]$ClientSecret,
        [Parameter(Mandatory)]
        [string]$FromEmail,
        [Parameter(Mandatory)]
        [string]$ToEmail,
        [Parameter(Mandatory)]
        [string]$Subject,
        [Parameter(Mandatory)]
        [string]$Body,
        [switch]$IsHtml,
        [Parameter()]
        [array]$Attachments
    )

    # Get access token
    $token = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

    # Initialize attachments array in payload
    $attachmentsPayload = @()

    # Process each attachment, if any
    if ($Attachments) {
        foreach ($attachmentPath in $Attachments) {
            if (Test-Path $attachmentPath) {
                # Read the file and convert to Base64
                $fileBytes = [System.IO.File]::ReadAllBytes($attachmentPath)
                $fileContent = [Convert]::ToBase64String($fileBytes)
                $fileName = [System.IO.Path]::GetFileName($attachmentPath)

                # Create an attachment object
                $attachmentObject = @{
                    "@odata.type" = "#microsoft.graph.fileAttachment"
                    name          = $fileName
                    contentBytes  = $fileContent
                    contentType   = "application/octet-stream"
                }
                $attachmentsPayload += $attachmentObject
            } else {
                Write-Warning "File not found: $attachmentPath"
            }
        }
    }

    # Build the email payload
    $emailPayload = @{
        message = @{
            subject = $Subject
            body = @{
                contentType = if ($IsHtml) { "HTML" } else { "Text" }
                content     = $Body
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $ToEmail
                    }
                }
            )
            from = @{
                emailAddress = @{
                    address = $FromEmail
                }
            }
            attachments = $attachmentsPayload
        }
    }

    # Send the email
    $response = Invoke-RestMethod -Uri "$global:GraphAPIUrl/users/$FromEmail/sendMail" `
                                  -Method POST `
                                  -Headers @{ Authorization = "Bearer $token" } `
                                  -Body (ConvertTo-Json $emailPayload -Depth 5) `
                                  -ContentType "application/json"

    if ([string]::IsNullOrEmpty($response)) {
        Write-Output "Email sent successfully to $ToEmail."
    } else {
        Write-Output "Failed to send email."
    }
}

# Export the functions
Export-ModuleMember -Function Send-Mail
