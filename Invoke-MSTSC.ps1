#TODO paramsets pro rezimy ktere nemohou byt spolu pouzity
#TODO nahazovani nazvu serveru z AD přes TAB
function Invoke-MSTSC {
    <#
    .SYNOPSIS
    Function for automatization of RDP connection to computer.
    By default it tries to read LAPS password and use it for connection (using cmdkey tool, that imports such credentials to Credential Manager temporarily). But can also be used for autofill of domain credentials (using AutoIt PowerShell module).

    .DESCRIPTION
    Function for automatization of RDP connection to computer.
    By default it tries to read LAPS password and use it for connection (using cmdkey tool, that imports such credentials to Credential Manager temporarily). But can also be used for autofill of domain credentials (using AutoIt PowerShell module).

    It has to be run from PowerShell console, that is running under account with permission for reading LAPS password!

    It uses LAPS official module for getting LAPS password and AutoItx PowerShell module for automatic filling of credentials into mstsc.exe app for RDP, in case LAPS password wasn't retrieved or domain account is used.

    It is working only on English OS.

    .PARAMETER computerName
    Name of remote computer/s

    .PARAMETER pickFreeComputer
    Switch for picking one of online computers stored in computerName parameter

    .PARAMETER useDomainAdminAccount
    Instead of local admin account, your adm_<username> domain account will be used.

    .PARAMETER credential
    Object with credentials, which should be used to authenticate to remote computer

    .PARAMETER port
    RDP port. Default is 3389

    .PARAMETER admin
    Switch. Use admin RDP mode

    .PARAMETER restrictedAdmin
    Switch. Use restrictedAdmin mode

    .PARAMETER remoteGuard
    Switch. Use remoteGuard mode

    .PARAMETER multiMon
    Switch. Use multiMon

    .PARAMETER fullScreen
    Switch. Open in fullscreen

    .PARAMETER public
    Switch. Use public mode

    .PARAMETER width
    Width of window

    .PARAMETER height
    Heigh of windows

    .PARAMETER gateway
    What gateway to use

    .PARAMETER localAdmin
    What is the name of local administrator, that will be used for LAPS conection

    .EXAMPLE
    Invoke-MSTSC pc1

    Run remote connection to pc1 using builtin administrator account and his LAPS password.

    .EXAMPLE
    Invoke-MSTSC pc1 -useDomainAdminAccount

    Run remote connection to pc1 using adm_<username> domain account.

    .EXAMPLE
    Invoke-MSTSC $somecomputers -pickFreeComputer

    Get online computer from $somecommputers list and RDP to it, using local admin account and LAPS password.

    .EXAMPLE
    $credentials = Get-Credential
    Invoke-MSTSC pc1 -credential $credentials

    Run remote connection to pc1 using credentials stored in $credentials

    .NOTES
    Automatic filling is working only on english operating systems.
    Author: Ondřej Šebela - ztrhgf@seznam.cz
    #>

    [CmdletBinding()]
    [Alias("rdp")]
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        $computerName
        ,
        [switch] $pickFreeComputer
        ,
        [switch] $useDomainAdminAccount
        ,
        [PSCredential] $credential
        ,
        [int] $port = 3389
        ,
        [switch] $admin
        ,
        [switch] $restrictedAdmin
        ,
        [switch] $remoteGuard
        ,
        [switch] $multiMon
        ,
        [switch] $fullScreen
        ,
        [switch] $public
        ,
        [int] $width
        ,
        [int] $height
        ,
        [string] $gateway
        ,
        [string] $localAdmin = "Administrator"
    )

    begin {
        # remove validation ValidateNotNullOrEmpty
        (Get-Variable computerName).Attributes.Clear()

        try {
            Write-Verbose "Get list of domain DCs"
            $DC = [System.Directoryservices.Activedirectory.Domain]::GetCurrentDomain().DomainControllers | ForEach-Object { ($_.name -split "\.")[0] }
        } catch {
            throw "Unable to contact your AD domain"
        }

        Write-Verbose "Get NETBIOS domain name"
        if (!$domainNetbiosName) {
            $domainNetbiosName = $env:userdomain

            if ($domainNetbiosName -eq $env:computername) {
                # function is running under local account therefore $env:userdomain cannot be used
                $domainNetbiosName = (Get-CimInstance Win32_NTDomain).DomainName # slow but gets the correct value
            }
        }
        Write-Verbose "Get domain name"
        if (!$domainName) {
            $domainName = (Get-CimInstance Win32_ComputerSystem).Domain
        }

        if ($pickFreeComputer) {
            try {
                $null = Get-Command Test-Connection2 -ErrorAction Stop
            } catch {
                throw "Command Test-Connection2 isn't available"
            }
        }

        $defaultRDP = Join-Path $env:USERPROFILE "Documents\Default.rdp"
        if (Test-Path $defaultRDP -ErrorAction SilentlyContinue) {
            Write-Verbose "RDP settings from $defaultRDP will be used"
        }

        if ($pickFreeComputer) {
            if ($computerName.GetType().name -ne 'string') {
                # more target computers
                # filter out offline computers
                $online = $computerName | ? { Test-Connection2 -Computername $_ }
                if ($online.count -gt 1) {
                    Write-Verbose "Just $online is online"
                    $name = Get-Random
                    # quser shows who is logged on
                    $dumb = Invoke-Command -ComputerName $online -AsJob -JobName $name -ScriptBlock {
                        $ErrorActionPreference = 'stop'
                        try {
                            # if there is active session, it means there is somebody logged on already
                            $sessions = quser.exe | Select-Object -Skip 1 | ForEach-Object {
                                $CurrentLine = $_.Trim() -Replace '\s+', ' ' -Split '\s'
                                $HashProps = @{
                                    UserName     = $CurrentLine[0]
                                    ComputerName = $env:COMPUTERNAME
                                }

                                # in case of disconnected session, output look little bit different
                                if ($CurrentLine[2] -eq 'Disc') {
                                    $HashProps.SessionName = $null
                                    $HashProps.Id = $CurrentLine[1]
                                    $HashProps.State = $CurrentLine[2]
                                    $HashProps.IdleTime = $CurrentLine[3]
                                    $HashProps.LogonTime = $CurrentLine[4..6] -join ' '
                                } else {
                                    $HashProps.SessionName = $CurrentLine[1]
                                    $HashProps.Id = $CurrentLine[2]
                                    $HashProps.State = $CurrentLine[3]
                                    $HashProps.IdleTime = $CurrentLine[4]
                                    $HashProps.LogonTime = $CurrentLine[5..7] -join ' '
                                }

                                $obj = New-Object -TypeName PSCustomObject -Property $HashProps | Select-Object -Property UserName, ComputerName, SessionName, Id, State, IdleTime, LogonTime
                                $obj
                            }

                            # if all sessions are in disconnected state it is ok to connect
                            if (($sessions | Where-Object { $_.state -eq 'Disc' }).count -eq $sessions.count) {
                                return $env:COMPUTERNAME
                            }

                        } catch {
                            # quser ends with error if no one is logged on
                            return $env:COMPUTERNAME
                        }
                    }

                    # get some free computer
                    $count = 0
                    while (1) {
                        ++$count
                        $result = Get-Job -IncludeChildJob -Name $name # -HasMoreData:$true -ChildJobState Completed | Receive-Job

                        if ($cResult = $result | Where-Object { $_.state -eq 'Completed' -and $_.hasmoredata -eq $true } | Receive-Job) {
                            # some tasks are finished pick some computer from them
                            $computerName = Get-Random -InputObject $cResult -Count 1
                            Get-Job -Name $name | Remove-Job -Force
                            break
                        } elseif (((($result | Where-Object { $_.state -ne 'Running' -and $_.hasmoredata -eq $false }).count - 1) -eq $online.count) -or $count -gt 100) {
                            # all jobs are finished, but there is no free computer
                            Get-Job -Name $name | Remove-Job -Force
                            $computerName = ''
                            break
                        }

                        Start-Sleep -Milliseconds 100
                    }
                } elseif ($online.count -eq 1) {
                    # just one computer is online
                    $computerName = $online
                } else {
                    Write-Warning "Not even one computer is online"
                    break
                }

                if (!$computerName) {
                    Write-Warning "Not even one computer is free"
                    break
                }
            } else {
                Write-Warning "You've entered just one computer, I will use it whether it is free or not"
            }
        }

        if ($computerName.GetType().name -ne 'string' -and !$pickFreeComputer) {
            while ($choice -notmatch "[Y|N]") {
                $choice = Read-Host "Do you really want to connect to all these computers:($($computerName.count))? (Y|N)"
            }
            if ($choice -eq "N") {
                break
            }
        }

        if ($credential) {
            $UserName = $Credential.UserName
            $Password = $Credential.GetNetworkCredential().Password
        } elseif ($useTierAccount) {
            # tier of the computer will define login I have to use
            # it is per machine, co I will fill it later
        } elseif ($useDomainAdminAccount) {
            if ($env:USERNAME -notmatch "^adm_") {
                $dAdmin = "adm_" + $env:USERNAME
            } else {
                $dAdmin = $env:USERNAME
            }

            $userName = "$domainNetbiosName\$dAdmin"
        } else {
            # no credentials were given, try to get LAPS password
            ++$tryLaps
        }

        # set MSTSC parameters
        switch ($true) {
            { $admin } { $mstscArguments += '/admin ' }
            { $restrictedAdmin } { $mstscArguments += '/restrictedAdmin ' }
            { $remoteGuard } { $mstscArguments += '/remoteGuard ' }
            { $multiMon } { $mstscArguments += '/multimon ' }
            { $fullScreen } { $mstscArguments += '/f ' }
            { $public } { $mstscArguments += '/public ' }
            { $width } { $mstscArguments += "/w:$width " }
            { $height } { $mstscArguments += "/h:$height " }
            { $gateway } { $mstscArguments += "/g:$gateway " }
        }

        $params = @{
            filePath = "$($env:SystemRoot)\System32\mstsc.exe"
        }

        if ($mstscArguments) {
            $params.argumentList = $mstscArguments
        }
    }

    process {
        foreach ($computer in $computerName) {
            # get just hostname
            if ($computer -match "\d+\.\d+\.\d+\.\d+") {
                # it is IP
                $computerHostname = $computer
            } else {
                # it is hostname or fqdn
                $computerHostname = $computer.split('\.')[0]
            }
            $computerHostname = $computerHostname.ToLower()

            if ($tryLaps -and $computerHostname -notin $DC.ToLower()) {
                Write-Verbose "Getting LAPS password for $computerHostname"
                $password = Get-LapsADPassword -Identity $computerName -AsPlainText | select -ExpandProperty Password

                if (!$password) {
                    Write-Warning "Unable to get LAPS password for $computerHostname."
                }
            }

            if ($tryLaps) {
                if ($computerHostname -in $DC.ToLower()) {
                    # connecting to DC (there are no local accounts
                    $userName = "$domainNetbiosName\$Env:USERNAME"
                } else {
                    # connecting to non-DC computer
                    if ($computerName -notmatch "\d+\.\d+\.\d+\.\d+") {
                        $userName = "$computerHostname\$localAdmin"
                    } else {
                        # IP was used instead of hostname, therefore I assume there is no LAPS
                        $UserName = " "
                    }
                }
            }

            # if hostname is not in FQDN and it is a server, I will add domain suffix (because of RDP certificate that is probably generated there)
            if ($computer -notmatch "\.") {
                Write-Verbose "Adding $domainName suffix to $computer"
                $computer = $computer + "." + $domainName
            }

            $connectTo = $computer

            if ($port -ne 3389) {
                $connectTo += ":$port"
            }

            # clone mstsc parameters just in case I am connecting to more than one computer, to be able to easily add /v hostname parameter
            $fParams = $params.Clone()

            #
            # log on automatization
            if ($password) {
                # I have password, so I will use cmdkey to store it in Cred. Manager
                Write-Verbose "Saving credentials for $computer and $userName to CredMan"
                $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
                $Process = New-Object System.Diagnostics.Process
                $ProcessInfo.FileName = "$($env:SystemRoot)\system32\cmdkey.exe"
                $ProcessInfo.Arguments = "/generic:TERMSRV/$computer /user:$userName /pass:`"$password`""
                $ProcessInfo.RedirectStandardOutput = ".\NUL"

                $ProcessInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $Process.StartInfo = $ProcessInfo
                [void]$Process.Start()
                $null = $Process.WaitForExit()

                if ($Process.ExitCode -ne 0) {
                    throw "Unable to add credentials to Cred. Manageru, but just for sure, check it."
                }

                # remote computer
                $fParams.argumentList += "/v $connectTo"
            } else {
                # I don't have credentials, so I have to use AutoIt for log on automation

                Write-Verbose "I don't have credentials, so AutoIt will be used instead"

                if ([console]::CapsLock) {
                    $keyBoardObject = New-Object -ComObject WScript.Shell
                    $keyBoardObject.SendKeys("{CAPSLOCK}")
                    Write-Warning "CAPS LOCK was turned on, disabling"
                }

                $titleCred = "Windows Security"
                if (((Get-AU3WinHandle -Title $titleCred) -ne 0) -and $password) {
                    Write-Warning "There is opened window for entering credentials. It has to be closed or auto-fill of credentials will not work."
                    Write-Host 'Enter any key to continue' -NoNewline
                    $null = [Console]::ReadKey('?')
                }
            }

            #
            # running mstsc
            Write-Verbose "Running mstsc.exe with parameter: $($fParams.argumentList)"
            Start-Process @fParams

            if ($password) {
                # I have password, so cmdkey was used for automation
                # so I will now remove saved credentials from Cred. Manager
                Write-Verbose "Removing saved credentials from CredMan"
                Start-Sleep -Seconds 1.5
                $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
                $Process = New-Object System.Diagnostics.Process
                $ProcessInfo.FileName = "$($env:SystemRoot)\system32\cmdkey.exe"
                $ProcessInfo.Arguments = "/delete:TERMSRV/$computer"
                $ProcessInfo.RedirectStandardOutput = ".\NUL"
                $ProcessInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $Process.StartInfo = $ProcessInfo
                [void]$Process.Start()
                $null = $Process.WaitForExit()

                if ($Process.ExitCode -ne 0) {
                    throw "Removal of credentials failed. Remove them manually from  Cred. Manager!"
                }
            } else {
                # I don't have password, so AutoIt will be used

                Write-Verbose "Automating log on process using AutoIt"

                try {
                    $null = Get-Command Show-AU3WinActivate -ErrorAction Stop
                } catch {
                    try {
                        $null = Import-Module AutoItX -ErrorAction Stop -Verbose:$false
                    } catch {
                        throw "Module AutoItX isn't available"
                    }
                }

                # click on "Show options" in mstsc console
                $title = "Remote Desktop Connection"
                Start-Sleep -Milliseconds 300 # to get the handle on last started mstsc
                $null = Wait-AU3Win -Title $title -Timeout 1
                $winHandle = Get-AU3WinHandle -Title $title
                $null = Show-AU3WinActivate -WinHandle $winHandle
                $controlHandle = Get-AU3ControlHandle -WinHandle $winhandle -Control "ToolbarWindow321"
                $null = Invoke-AU3ControlClick -WinHandle $winHandle -ControlHandle $controlHandle
                Start-Sleep -Milliseconds 600

                # fill computer and username
                Write-Verbose "Connecting to: $connectTo as: $userName"
                Send-AU3Key -Key "{CTRLDOWN}A{CTRLUP}{DELETE}" # delete any existing text
                Send-AU3Key -Key "$connectTo{DELETE}" # delete any suffix, that could be autofilled there

                Send-AU3Key -Key "{TAB}"
                Start-Sleep -Milliseconds 400

                Send-AU3Key -Key "{CTRLDOWN}A{CTRLUP}{DELETE}" # delete any existing text
                Send-AU3Key -Key $userName
                Send-AU3Key -Key "{ENTER}"
            }

            # # accept any untrusted certificate
            # $title = "Remote Desktop Connection"
            # $null = Wait-AU3Win -Title $title -Timeout 1
            # $winHandle = ''
            # $count = 0
            # while ((!$winHandle -or $winHandle -eq 0) -and $count -le 40) {
            #     # nema smysl cekat moc dlouho, protoze certak muze byt ok nebo uz ma vyjimku
            #     $winHandle = Get-AU3WinHandle -Title $title -Text "The certificate is not from a trusted certifying authority"
            #     Start-Sleep -Milliseconds 100
            #     ++$count
            # }
            # # je potreba potvrdit nesedici certifikat
            # if ($winHandle) {
            #     $null = Show-AU3WinActivate -WinHandle $winHandle
            #     Start-Sleep -Milliseconds 100
            #     $controlHandle = Get-AU3ControlHandle -WinHandle $winhandle -Control "Button5"
            #     $null = Invoke-AU3ControlClick -WinHandle $winHandle -ControlHandle $controlHandle
            # }
        }
    }
}
