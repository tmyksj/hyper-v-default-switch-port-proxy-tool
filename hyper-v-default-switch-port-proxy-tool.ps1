$AppName = "Hyper-V Default Switch Port Proxy Tool"

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-Command", "if ((Get-ExecutionPolicy) -ne 'AllSigned') { Set-ExecutionPolicy -Scope Process RemoteSigned } &'$PSCommandPath'" -Verb runas
    Exit
}

:loop while ($TRUE) {
    $command = (Read-Host $AppName).Split(" ")

    switch ($command[0]) {
        "" {
        }
        "add" {
            if (($command.Length -ne 3) -And ($command.Length -ne 4)) {
                Write-Host "illegal argument."
                Write-Host
                continue loop
            }

            $listenPort = $command[1]
            $connectAddress = $command[2]
            $connectPort = if ($command.Length -eq 3) { $command[1] } else { $command[3] }

            netsh interface portproxy add v4tov4 listenaddress=0.0.0.0 listenport=$listenPort connectaddress=$connectAddress connectport=$connectPort
            netsh advfirewall firewall add rule name="$AppName $listenPort" dir=in action=allow localport=$listenPort protocol=tcp
        }
        "delete" {
            if ($command.Length -ne 2) {
                Write-Host "illegal argument."
                Write-Host
                continue loop
            }

            $listenPort = $command[1]

            netsh advfirewall firewall delete rule name="$AppName $listenPort"
            netsh interface portproxy delete v4tov4 listenaddress=0.0.0.0 listenport=$listenPort
        }
        "exit" {
            break loop
        }
        "help" {
            Write-Host
            Write-Host "command list:"
            Write-Host "    - add {listen port} {destination vm name} [{destination vm port}]"
            Write-Host "    - delete {listen port}"
            Write-Host "    - exit"
            Write-Host "    - help"
            Write-Host "    - show"
            Write-Host
        }
        "show" {
            netsh interface portproxy show all
            netsh advfirewall firewall show rule name=all dir=in | Select-String $AppName
            Write-Host
        }
        default {
            Write-Host "unknown command."
            Write-Host
        }
    }
}
