PARAM ($UPNs = [string], [switch]$RevokeAllUsers)
#************************************************
# RevokeRefreshTokens.ps1
# Version 1.0
# Date: 3-30-2018
# Author: Tim Springston [MSFT] 
# Description: This script will revoke the Azure AD user refresh tokens for a specificed user or users.
# More info at https://docs.microsoft.com/en-us/powershell/module/azuread/revoke-azureaduserallrefreshtoken?view=azureadps-2.0
#************************************************
cls

$ErrorActionPreference = 'SilentlyContinue'

function RevokeRefreshTokensforSingleUser 
 { PARAM ($UPN) 
          $User = Get-AzureADUser -objectID $UPN
          try {
                Revoke-AzureADUserAllRefreshToken -ObjectId $User.objectid
                $User = Get-AzureADUser -objectID $User.objectid 
                Write-Host "Refresh tokens for user" $User.UserPrincipalName "were succesfully revoked. Refresh tokens issued after"  $User.RefreshTokensValidFromDateTime "will be valid." -ForegroundColor Green
                }
                catch {Write-Host "The user name" $User "is not in the tenant or the user principal name is not in the correct format." -ForegroundColor Yellow}
}

function RevokeRefreshTokensforMultipleUsers
 { PARAM ($Users) 
          
          foreach ($User in $Users)
            {
            $UserObject = Get-AzureADUser -objectID $User
            try {
                Revoke-AzureADUserAllRefreshToken -ObjectId $UserObject.objectid 
                $UserObject = Get-AzureADUser -objectID $UserObject.objectid  
                Write-Host "Refresh tokens for user" $UserObject.UserPrincipalName "were succesfully revoked. Refresh tokens issued after"  $UserObject.RefreshTokensValidFromDateTime "will be valid." -ForegroundColor Green
                }
                catch {Write-Host "The user name" $User "is not in the tenant or the user principal name is not in the correct format." -ForegroundColor Yellow}
            $UserObject = $null
            }
}

function RevokeRefreshTokensforAllUsers
 {
          $AllUsers = Get-AzureADUser -All $true
          $ResultsFile = $pwd.path + '\RefreshTokenRevocationResults.txt'
          foreach ($User in $AllUsers)
                {
                try {
                    Revoke-AzureADUserAllRefreshToken -ObjectId $User.objectid
                    $User = Get-AzureADUser -objectID $User.objectid
                    $SuccesString = "Refresh tokens for user " + $User.UserPrincipalName + " were succesfully revoked. Refresh tokens issued after " + $User.RefreshTokensValidFromDateTime + " will be valid."
                    Write-Host $SuccesString -ForegroundColor Green
                    $SuccesString | Out-File -Append -Filepath $ResultsFile
                    }
                catch {
                      $FailString =  "The user name "+ $User.UserPrincipalName + " is not in the tenant or the user principal name is not in the correct format."
                      Write-Host $FailString -ForegroundColor Yellow
                      $FailString | Out-File -Append -FilePath $ResultsFile
                      }
                }
           Write-host "Revocation results file is located at $ResultsFile."

}


if ($UPNs -match '\w+|\d@\w+|d\.\w+') #minimal qualification of UPN format 
   {
    if (!(Get-Module -Name "AzureADPreview"))
          {
          Write-host "The AzureADPreview Module is not installed. Installing now..."
          Install-Module -Name "AzureADPreview" -AllowClobber -Force -Confirm
          }
        Import-Module AzureADPreview
        $Connection = Connect-AzureAD -Credential (Get-Credential) -InformationAction SilentlyContinue
    
    if (!($RevokeAllUsers))
        {
        if ($UPNs.count -eq 1) {RevokeRefreshTokensforSingleUser -UPN $UPNs}
        if ($UPNs.count -ge 2) {RevokeRefreshTokensforMultipleUsers -Users $UPNs}
        }
        else {RevokeRefreshTokensforAllUsers}

    }
          else  {Write-Host "The user name" $UPN "is not in the tenant or the user principal name is not in the correct format."}



