function Send-LAPSPassword {
    <#
    .SYNOPSIS
    Function retrieve LAPS password for defined computer and send it to specified email address (using SSL)
    Automatically reset LAPS password after specified amount of time.

    .DESCRIPTION
    Function retrieve LAPS password for defined computer and send it to specified email address 
    Automatically reset LAPS password after specified amount of time.
    
    It is essential to run function under account with permissions to read LAPS password from Active Directory!

    .PARAMETER computerName
    Computer name for LAPS retrieval.

    .PARAMETER emailTo
    To what email address send email with LAPS password.

    .PARAMETER from
    From what address should script send email.
    
    .PARAMETER smtpServer
    What smtp server should it use.

    .PARAMETER resetAfterMinutes
    Number of minutes to wait before reset of LAPS password occur.

    Default is 60 minutes.

    .EXAMPLE
    Send-LAPSPassword -computerName somepc -emailTo admin@domain.cz

    Send LAPS password for computer somepc to address admin@domain.cz. After 60 minutes, the password will be will reset.
 #>   

    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
                If (((New-Object DirectoryServices.DirectorySearcher -Property @{Filter = "(&(objectCategory=computer)(name=*$_*))"; PageSize = 500}).findall()) -ne '' ) {
                    $true
                } else {
                    Throw "Computer $_ doesn't exist in domain"
                }
            })]
        [string] $computerName
        ,
        [Parameter(Position = 1, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string] $emailTo
        ,
        [Parameter(Mandatory = $True)]        
        [string] $from
        ,
        [Parameter(Mandatory = $True)]        
        [string] $smtpServer
        ,
        [Parameter(Position = 2)]
        [int] $resetAfterMinutes = 60
    )

    begin {
        try {
            $null = import-module AdmPwd.PS -ErrorAction Stop
        } catch {
            throw "Module AdmPwd.PS is not available. Import it first!"
        }

        try {
            $null = Get-Command Get-AdmPwdPassword, Reset-AdmPwdPassword -ErrorAction Stop
        } catch {
            throw "Some cmdlets are missing: Get-AdmPwdPassword or Reset-AdmPwdPassword"
        }
    }

    process {
        $password = (Get-AdmPwdPassword $computerName).password

        if (!$password) {
            throw "Retrieval of LAPS password for  $computerName wasn't succesfull. Does it use LAPS?"
        }

        # for security reasons send in email message just password (not computer name or that it is password)
        # optionally you can attach some 3 random chars to end of it, for case, that email message could be compromised, so you just have to ignore last 3 chars when you use it :)
        # $body = $Body + $(-join ((65..90) + (97..122) | Get-Random -Count 3 | % {[char]$_}))
        "Password for $computerName is going to be send to $emailTo. After $resetAfterMinutes minutes it will be automatically reset for security reasons."
        Send-MailMessage -To $emailTo -from $from -Subject 'response' -Body $password -SmtpServer $smtpServer -UseSsl

        # reset LAPS password after x minutes (for security reasons)
        $null = Start-Job -Name "resPass$($computerName.ToUpper())" {
            param ($resetAfterMinutes, $computerName)
            Start-Sleep ($resetAfterMinutes * 60)
            Reset-AdmPwdPassword $computerName
            Invoke-Command $computerName {gpupdate /force}
        } -ArgumentList $resetAfterMinutes, $computerName
    }
}