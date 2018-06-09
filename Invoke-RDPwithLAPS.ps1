function Invoke-RDPwithLAPS {
    <#
    .SYNOPSIS
    Function for automatization of RDP connection to computer, using builtin administrator account and LAPS password.
    It uses AdmPwd.PS and AutoItx powershell modules for getting LAPS password and automatic fill into mstsc.exe app for RDP. 

    .DESCRIPTION
    Function for automatization of RDP connection to computer, using builtin administrator account and LAPS password
    It has to be run from powershell console, that is running under account with permission for reading LAPS password!
    Function runs mstsc.exe and automatically fills in computername, username and LAPS password. 
    It is working only on English OS.

    .PARAMETER computerName
    Name of remote computer/s

    .PARAMETER pickFreeComputer
    Switch for picking one of online computers stored in computerName parameter 

    .PARAMETER credential
    Object with credentials, which should be used to authenticate to remote computer

    .PARAMETER port
    RDP port. Default is 3389

    .PARAMETER admin
    Switch. Use admin mode

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

    .EXAMPLE
    Invoke-RDPwithLAPS pc1

    Run remote connection to pc1 using builtin administrator account and his LAPS password.

    .EXAMPLE
    Invoke-RDPwithLAPS $somecomputers -pickFreeComputer

    Get online computer from $somecommputers list and RDP to it, using local admin account and LAPS password.

    .EXAMPLE
    $credentials = Get-Credential
    Invoke-RDPwithLAPS pc1 -credential $credentials 

    RUn remote connection to pc1 using credentials stored in $credentials

    .NOTES
    Automatic filling is working only on english operating systems. 
    #>

    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
                If ($_.GetType().name -ne 'string' -or ($_.GetType().name -eq 'string' -and (((New-Object DirectoryServices.DirectorySearcher -Property @{Filter = "(&(objectCategory=computer)(name=*$_*))"; PageSize = 500}).findall()) -ne '' ) -or $_ -match "\d+\.\d+\.\d+\.\d+")) {
                    $true
                } else {
                    Throw "Stroj $_ v domene neexistuje"
                }
            })]
        $computerName
        ,
        [switch] $pickFreeComputer
        ,
        [switch] $useTierAccount
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
    )

    begin {
        try {
            $null = Import-Module AutoItX -ErrorAction Stop -verbose:$false
        } catch {
            throw "Module AutoItX is not available. Import it first!"
        }

        try {
            $null = import-module AdmPwd.PS -ErrorAction Stop -verbose:$false
        } catch {
            throw "Module AdmPwd.PS is not available. Import it first!"
        }

        if ($pickFreeComputer) {
            # super fast ping function
            Function Test-Connection2 {
                <#
                    .SYNOPSIS
                        Funkce k otestovani dostupnosti stroju. 
            
                    .DESCRIPTION
                        Funkce k otestovani dostupnosti stroju. Pouziva asynchronni ping. 
            
                    .PARAMETER Computername
                        List of computers to test connection
            
                    .PARAMETER DetailedTest
                        Prepinac. Pomalejsi metoda testovani vyzadujici modul psasync. 
                        Krome pingu otestuje i dostupnost c$ sdileni a RPC.
            
                        Aby melo smysl, je potreba mit na danych strojich prava pro pristup k c$ sdileni!
            
                    .PARAMETER Repeat
                        Prepinac. Donekonecna bude pingat vybrane stroje.
                        Neda se pouzit spolu s DetailedTest
            
                    .PARAMETER JustResponding
                        Vypise jen stroje, ktere odpovidaji
            
                    .PARAMETER JustNotResponding
                        Vypise jen stroje, ktere neodpovidaji
            
                    .PARAMETER Timeout
                        Timeout in milliseconds
            
                    .PARAMETER TimeToLive
                        Sets a time to live on ping request
            
                    .PARAMETER Fragment
                        Tells whether to fragment the request
            
                    .PARAMETER Buffer
                        Supply a byte buffer in request
            
                    .NOTES
                        Vychazi ze skriptu Test-ConnectionAsync od Boe Prox
            
                    .EXAMPLE
                        Test-Connection2 -Computername server1,server2,server3
            
                        Computername                Result
                        ------------                ------
                        Server1                     Success
                        Server2                     TimedOut
                        Server3                     No such host is known
                    
                    .EXAMPLE
                        $offlineStroje = Test-Connection2 -Computername server1,server2,server3 -JustNotResponding
            
                    .EXAMPLE
                        if (Test-Connection2 bumpkin -JustResponding) {"Bumpkin bezi"}
                #>
            
                [OutputType('Net.AsyncPingResult')]
                [cmdletbinding(DefaultParameterSetName = 'Default')]
                Param (
                    [Parameter(Mandatory = $true, Position = 0, ValueFromPipelinebyPropertyName = $true, ValueFromPipeline = $true)]
                    [string[]] $Computername
                    ,
                    [Parameter(Mandatory = $false, ParameterSetName = "Online")] 
                    [switch] $JustResponding
                    ,
                    [Parameter(Mandatory = $false, ParameterSetName = "Offline")] 
                    [switch] $JustNotResponding
                    ,
                    [switch] $DetailedTest
                    ,
                    [Alias('t')]
                    [switch] $Repeat
                    ,
                    [parameter()]
                    [int32] $Timeout = 100
                    ,
                    [parameter()]
                    [Alias('Ttl')]
                    [int32] $TimeToLive = 128
                    ,
                    [parameter()]
                    [switch] $Fragment
                    ,
                    [parameter()]
                    [byte[]] $Buffer
                )
            
                Begin {
                    if ($DetailedTest -and $Repeat) {
                        Write-Warning "Prepinac detailed, se neda pouzit v kombinaci s repeat."
                        $DetailedTest = $false
                    }
            
                    if ($DetailedTest) {
                        if (! (Get-Module psasync)) {
                            throw "Pro detailni otestovani dostupnosti je potreba psasync modul"
                        }
                        
                        $AsyncPipelines = @()
                        $pool = Get-RunspacePool 30
                        $scriptblock = `
                        {
                            param($computer, $JustResponding, $JustNotResponding)
                            # vytvorim si objekt s atributy
                            $Object = [pscustomobject] @{
                                ComputerName = $computer
                                Result       = ""
                            }
            
                            if (Test-Connection $computer -count 1 -quiet) {
                                if (! (Get-WmiObject win32_computersystem -ComputerName $Computer -ErrorAction SilentlyContinue)) {
                                    $Object.Result = "RPC not available"
                                } elseif (Test-Path \\$computer\c$) { 
                                    $Object.Result = "Success"
                                } else { 
                                    $Object.Result = "c$ share not available"
                                }     
                            } else { 
                                $Object.Result = "TimedOut"
                            }
                            
                            if (($JustResponding -and $Object.Result -eq 'Success') -or ($JustNotResponding -and $Object.Result -ne 'Success')) {
                                $Object.ComputerName
                            } elseif (!$JustResponding -and !$JustNotResponding) {
                                $Object
                            }
                        }
                    } else {
                        If (-NOT $PSBoundParameters.ContainsKey('Buffer')) {
                            $Buffer = 0x61, 0x62, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6a, 0x6b, 0x6c, 0x6d, 0x6e, 0x6f, 
                            0x70, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76, 0x77, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69
                        }
                        $PingOptions = New-Object System.Net.NetworkInformation.PingOptions
                        $PingOptions.Ttl = $TimeToLive
                        If (-NOT $PSBoundParameters.ContainsKey('Fragment')) {
                            $Fragment = $False
                        }
                        $PingOptions.DontFragment = $Fragment
                    }
                }
            
                Process {
                    if ($DetailedTest) {
                        foreach ($computer in $ComputerName) {
                            $AsyncPipelines += Invoke-Async -RunspacePool $pool -ScriptBlock $ScriptBlock -Parameters $computer, $JustResponding, $JustNotResponding
                        }
                    } 
                }
            
                End {
                    if ($DetailedTest) {
                        Receive-AsyncResults -Pipelines $AsyncPipelines -ShowProgress
                    } else {
                        while (1) {
                            $Task = ForEach ($Computer in $Computername) {
                                [pscustomobject] @{
                                    ComputerName = $Computer
                                    Task         = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer, $Timeout, $Buffer, $PingOptions)
                                }
                            }   
            
                            Try {
                                [void][Threading.Tasks.Task]::WaitAll($Task.Task)
                            } Catch {}
            
                            $Task | ForEach {
                                If ($_.Task.IsFaulted) {
                                    $Result = $_.Task.Exception.InnerException.InnerException.Message
                                    $IPAddress = $Null
                                } Else {
                                    $Result = $_.Task.Result.Status
                                    $IPAddress = $_.task.Result.Address.ToString()
                                }
            
                                $Object = [pscustomobject]@{
                                    ComputerName = $_.Computername
                                    #IPAddress = $IPAddress
                                    Result       = $Result
                                }
            
                                $Object.pstypenames.insert(0, 'Net.AsyncPingResult')
            
                                if (($JustResponding -and $Object.Result -eq 'Success') -or ($JustNotResponding -and $Object.Result -ne 'Success')) {
                                    $Object.ComputerName
                                } elseif (!$JustResponding -and !$JustNotResponding) {
                                    $Object
                                }
                            }
            
                            # ukoncim while cyklus pokud neni receno, ze se maji vysledky neustale vypisovat
                            if (!$Repeat) {
                                break
                            } else {
                                sleep 1 #TODO pri prvnim vypsani vysledku se vypise az po tomto sleepu, proto tu nemam vetsi hodnotu, vyresit
                            }
                        } # end while
                    }
                }
            }
        }

        if ((Get-Host).CurrentCulture.name -ne 'en-US' -or (Get-Host).CurrentUICulture.name -ne 'en-US') {
            throw "Function works only on english OS. Exitting"
        }

        while ([console]::CapsLock) { 
            Write-Warning "CAPS LOCK is enabled. For continue turn it off" 
            sleep 2 
        } 

        $defaultRDP = Join-Path $env:USERPROFILE "Documents\Default.rdp"
        if (Test-Path $defaultRDP -ErrorAction SilentlyContinue) {
            Write-Verbose "Use RDP settings from $defaultRDP"
        }    

        if ($pickFreeComputer) {
            if ($computerName.GetType().name -ne 'string') {
                # computerName contains more computers
               
                # remove offline computers
                $online = Test-Connection2 $computerName -JustResponding
                if ($online.count -gt 1) {
                    $name = Get-Random
                    $dumb = Invoke-Command -comp $online -AsJob -jobname $name -scriptblock {
                        $ErrorActionPreference = 'stop'
                        try {
                            # quser end with error if nobody is logged on
                            $dumb = quser
                        } catch {
                            # nobody is logged on
                            return $env:COMPUTERNAME
                        }
                    }

                    # if some of running jobs is finished, i write out its output and end the others
                    while (1) {
                        $result = Get-Job -IncludeChildJob -Name $name -HasMoreData:$true -ChildJobState Completed | Receive-Job 
                        if ($result) {
                            $computerName = Get-Random -InputObject $result -Count 1
                            Get-Job -Name $name | Remove-Job -Force
                            break
                        }
                        Start-Sleep -Milliseconds 50
                    }
                } elseif ($online -eq 1) {
                    # just one computer is online
                    $computerName = $online
                } else {
                    Write-Warning "Neither of computers is online"
                    return
                }

                if (!$computerName) {
                    Write-Warning "Neither of computers is free now"
                    return
                }
            } else {
                Write-Warning "You entered just one computer, I try to connect to it without testing its status"
            }
        }
        

        if ($computerName.GetType().name -ne 'string' -and !$pickFreeComputer) {
            while ($choice -notmatch "[Y|N]") {
                $choice = read-host "Do you really want to connect to all of these computers: ($($computerName.count))? (Y|N)"
            }
            if ($choice -eq "N") {
                break
            }
        }
      

        if ($credential) {
            $UserName = $Credential.UserName
            $Password = $Credential.GetNetworkCredential().Password
        } else {
            # user didnt enter credentials, I will try to get LAPS passwor
            ++$tryLaps
            $userName = "administrator"
        }


        switch ($true) {
            {$admin} {$mstscArguments += '/admin '}
            {$restrictedAdmin} {$mstscArguments += '/restrictedAdmin '}
            {$remoteGuard} {$mstscArguments += '/remoteGuard '}
            {$multiMon} {$mstscArguments += '/multimon '}
            {$fullScreen} {$mstscArguments += '/f '}
            {$public} {$mstscArguments += '/public '}
            {$width} {$mstscArguments += "/w:$width "}
            {$height} {$mstscArguments += "/h:$height "}
            {$gateway} {$mstscArguments += "/g:$gateway "}
        }

        $params = @{
            filePath = 'mstsc.exe'
        }

        if ($mstscArguments) {
            $params.argumentList = $mstscArguments
        }
    }

    process {
        foreach ($computer in $computerName) {
            if ($computer -match "\d+\.\d+\.\d+\.\d+") {
                $computerHostname = $computer
            } else {
                $computerHostname = $computer.split('\.')[0]
            }
            if ($tryLaps -and $computerHostname -notin $DC) {
                $password = (Get-AdmPwdPassword $computerHostname).password

                if (!$password) {
                    Write-Warning "LAPS password retrieval failed for computer $computerHostname. Are you running this function in console with right permissions?"
                }
            }
            
            $connectTo = $computer
            if ($port -ne 3389) {
                $connectTo += ":$port"
            }

            $titleCred = "Windows Security" 
            if (((Get-AU3WinHandle -Title $titleCred) -ne 0) -and $password) {
                Write-Warning "You have somewhere opened 'Windows Security' dialog for credentials input. CLose it before proceeding."
                Write-Host 'Press any key for continue' -NoNewline
                $null = [Console]::ReadKey('?')
            }

            Write-Verbose "Parameters: $($params.argumentList)"
            Start-Process @params
        
            #
            # AUTOIT automatization of logon
            # 

            # opening "Show options" in mstsc console            
            $title = "Remote Desktop Connection"
            Start-Sleep -Milliseconds 300
            $null = Wait-AU3Win -Title $title -Timeout 1
            $winHandle = Get-AU3WinHandle -Title $title
            $null = Show-AU3WinActivate -WinHandle $winHandle
            $controlHandle = Get-AU3ControlHandle -WinHandle $winhandle -Control "ToolbarWindow321" 
            $null = Invoke-AU3ControlClick -WinHandle $winHandle -ControlHandle $controlHandle 
            Start-Sleep -Milliseconds 600


            # filling hostname and username
            Write-Verbose "Connecting to: $connectTo as: $userName"
            Send-AU3Key -Key "{CTRLDOWN}A{CTRLUP}{DELETE}"       
            Send-AU3Key -Key $connectTo

            Send-AU3Key -Key "{TAB}"
            Start-Sleep -Milliseconds 400

            Send-AU3Key -Key "{CTRLDOWN}A{CTRLUP}{DELETE}"        
            Send-AU3Key -Key "$userName"
            Send-AU3Key -Key "{ENTER}"

            
            # filling password
            $titleCred = "Windows Security" 
            $null = Wait-AU3Win -Title $titleCred -Timeout 1
            $winHandle = Get-AU3WinHandle -Title $titleCred
            $count = 0
            while ((!$winHandle -or $winHandle -eq 0) -and $count -le 20) {
                $winHandle = Get-AU3WinHandle -Title $titleCred
                Start-Sleep -Milliseconds 200
                ++$count
            }
            $null = Show-AU3WinActivate -WinHandle $winHandle
            if ($password) {
                Send-AU3Key -Key "$password" -mode 1 # send text as raw
                Send-AU3Key -Key "{ENTER}"
            }

            # accept untrusted certificate
            $title = "Remote Desktop Connection"
            $null = Wait-AU3Win -Title $title -Timeout 1
            $winHandle = ''
            $count = 0
            while ((!$winHandle -or $winHandle -eq 0) -and $count -le 10) {
                $winHandle = Get-AU3WinHandle -Title $title -Text "The certificate is not from a trusted certifying authority"
                Start-Sleep -Milliseconds 200
                ++$count
            }
            if ($winHandle) {
                $null = Show-AU3WinActivate -WinHandle $winHandle
                Start-Sleep -Milliseconds 300
                $controlHandle = Get-AU3ControlHandle -WinHandle $winhandle -Control "Button5"
                $null = Invoke-AU3ControlClick -WinHandle $winHandle -ControlHandle $controlHandle
            }
        }
    }
}