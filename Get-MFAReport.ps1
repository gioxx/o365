Function Get-PerUserMFAStatus {
<#
.SYNOPSIS
    Get Per-User MFA Status using MSOnline Powershell Module
 
.NOTES
    Name: Get-PerUserMFAStatus
    Author: theSysadminChannel
    Version: 1.0
    DateCreated: 2021-Feb-3
 
.LINK
    https://thesysadminchannel.com/get-per-user-mfa-status-using-powershell -
#>
 
    [CmdletBinding(DefaultParameterSetName='All')]
    param(
        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'UPN',
            Position = 0
        )]
        [string[]]  $UserPrincipalName,
 
 
        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'All'
        )]
        [switch]    $All
 
    )
 
    BEGIN {
        if (-not (Get-MsolDomain -ErrorAction SilentlyContinue)) {
            Write-Error "You must connect to the MSolService to continue" -ErrorAction Stop
        }
    }
 
    PROCESS {
        if ($PSBoundParameters.ContainsKey('UserPrincipalName')) {
            $MsolUserList = foreach ($MsolUser in $UserPrincipalName) {
                try {
                    Get-MsolUser -UserPrincipalName $MsolUser -ErrorAction Stop
                     
                } catch {
                    Write-Error $_.Exception.Message
 
                }
            }
        } else {
            $MsolUserList = Get-MsolUser -All -ErrorAction Stop | Where-Object {$_.UserType -ne 'Guest' -and $_.DisplayName -notmatch 'On-Premises Directory Synchronization'}
        }
 
        #Now that we have our UserList, lets check the per-user mfa status
        foreach ($User in $MsolUserList) {
            if ($User.StrongAuthenticationRequirements) {
                $PerUserMFAState = $User.StrongAuthenticationRequirements.State
 
              } else {
                $PerUserMFAState = 'Disabled'
 
            }
 
            $MethodType = $User.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq $true} | select -ExpandProperty MethodType
             
            if ($MethodType) {
                switch ($MethodType) {
                    'OneWaySMS'            {$DefaultMethodType = 'SMS Text Message'  }
                    'TwoWayVoiceMobile'    {$DefaultMethodType = 'Call to Phone'     }
                    'PhoneAppOTP'          {$DefaultMethodType = 'TOTP'              }
                    'PhoneAppNotification' {$DefaultMethodType = 'Authenticator App' }
                }
              } else {
                $DefaultMethodType = 'Not Enabled'
            }
     
            [PSCustomObject]@{
                UserPrincipalName    = $User.UserPrincipalName
                DisplayName          = $User.DisplayName
                PerUserMFAState      = $PerUserMFAState
                DefaultMethodType    = $DefaultMethodType
 
            }
 
            $MethodType        = $null
        }
 
    }
 
    END {}
}

Get-PerUserMFAStatus | Export-CSV ".\MFA.csv" -Delimiter ";" -NoTypeInformation