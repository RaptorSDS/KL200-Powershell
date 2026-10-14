function Send-Command {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [byte[]]$Command
    )
    $Checksum = 0
    foreach ($byte in $Command) {
        $Checksum = $Checksum -bxor $byte
    }
    $Command += $Checksum
    $SerialPort.Write($Command, 0, $Command.Length)
}

function Calculate-Checksum {
    param (
        [byte[]]$Data
    )
    $Checksum = 0
    foreach ($byte in $Data) {
        $Checksum = $Checksum -bxor $byte
    }
    return $Checksum
}

function Initialize-Sensor {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [int]$BaudRate
    )
    $SerialPort.BaudRate = $BaudRate
    $SerialPort.Open()
}

function Restore-Factory-Settings {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [bool]$HardReset
    )
    $ResetByte = if ($HardReset) { 0xFE } else { 0xFD }
    $Command = [byte[]](0x62, 0x39, 0x09, 0xFF, 0xFF, 0xFF, 0xFF, $ResetByte)
    Send-Command -SerialPort $SerialPort -Command $Command
}

function Change-Address {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [uint16]$Address
    )
    if ($Address -gt 0xFFFE) {
        Write-Host "Address out of range"
        return
    }
    $Command = [byte[]](0x62, 0x32, 0x09, 0xFF, 0xFF, ($Address -shr 8), ($Address -band 0xFF))
    Send-Command -SerialPort $SerialPort -Command $Command
}

function Change-Baud-Rate {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [byte]$BaudRate
    )
    if ($BaudRate -gt 9) {
        Write-Host "Baud rate out of range"
        return
    }
    $Command = [byte[]](0x62, 0x30, 0x09, 0xFF, 0xFF, $BaudRate)
    Send-Command -SerialPort $SerialPort -Command $Command
}

function Set-Upload-Mode {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [bool]$AutoUpload
    )
    $Mode = if ($AutoUpload) { 1 } else { 0 }
    $Command = [byte[]](0x62, 0x34, 0x09, 0xFF, 0xFF, $Mode)
    Send-Command -SerialPort $SerialPort -Command $Command
}

function Set-Upload-Interval {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [byte]$Interval
    )
    if ($Interval -lt 1 -or $Interval -gt 100) {
        Write-Host "Interval out of range"
        return
    }
    $Command = [byte[]](0x62, 0x35, 0x09, 0xFF, 0xFF, $Interval)
    Send-Command -SerialPort $SerialPort -Command $Command
}

function Set-LED-Mode {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [byte]$Mode
    )
    if ($Mode -gt 3) {
        Write-Host "LED mode out of range"
        return
    }
    $Command = [byte[]](0x62, 0x37, 0x09, 0xFF, 0xFF, $Mode)
    Send-Command -SerialPort $SerialPort -Command $Command
}

function Set-Relay-Mode {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [byte]$Mode
    )
    if ($Mode -gt 1) {
        Write-Host "Relay mode out of range"
        return
    }
    $Command = [byte[]](0x62, 0x38, 0x09, 0xFF, 0xFF, $Mode)
    Send-Command -SerialPort $SerialPort -Command $Command
}

function Set-Communication-Mode {
    param (
        [System.IO.Ports.SerialPort]$SerialPort,
        [byte]$Mode
    )
    if ($Mode -gt 1) {
        Write-Host "Communication mode out of range"
        return
    }
    $Command = [byte[]](0x62, 0x31, 0x09, 0xFF, 0xFF, $Mode)
    Send-Command -SerialPort $SerialPort -Command $Command
}

function Read-Distance {
    param (
        [System.IO.Ports.SerialPort]$SerialPort
    )
    $Command = [byte[]](0x62, 0x33, 0x09, 0xFF, 0xFF, 0x00, 0x00, 0x00)
    Send-Command -SerialPort $SerialPort -Command $Command

    Start-Sleep -Milliseconds 100

    $Buffer = New-Object byte[] 9
    $BytesRead = $SerialPort.Read($Buffer, 0, 9)
    if ($BytesRead -eq 9 -and $Buffer[0] -eq 0x62 -and $Buffer[1] -eq 0x33) {
        $Length = $Buffer[2]
        $Address = ($Buffer[3] -shl 8) -bor $Buffer[4]
        $RawDistance = ($Buffer[5] -shl 8) -bor $Buffer[6]
        $Checksum = $Buffer[8]
        $CalcChecksum = Calculate-Checksum($Buffer[0..7])

        if ($Checksum -eq $CalcChecksum) {
            return $RawDistance
        } else {
            Write-Host "Checksum mismatch"
            return $null
        }
    } else {
        Write-Host "Invalid response"
        return $null
    }
}

function Show-Menu {
    param (
        [System.IO.Ports.SerialPort]$SerialPort
    )
    while ($true) {
        Clear-Host
        Write-Host "XKC-KL200 Controller Menu"
        Write-Host "1. Restore Factory Settings"
        Write-Host "2. Change Address"
        Write-Host "3. Change Baud Rate"
        Write-Host "4. Set Upload Mode"
        Write-Host "5. Set Upload Interval"
        Write-Host "6. Set LED Mode"
        Write-Host "7. Set Relay Mode"
        Write-Host "8. Set Communication Mode"
        Write-Host "9. Read Distance"
        Write-Host "10. Change COM Port"
        Write-Host "11. Exit"
        $Choice = Read-Host "Select an option"

        switch ($Choice) {
            1 {
                $HardReset = Read-Host "Hard reset (true/false)"
                Restore-Factory-Settings -SerialPort $SerialPort -HardReset [bool]::Parse($HardReset)
            }
            2 {
                $Address = Read-Host "Enter new address (hex)"
                Change-Address -SerialPort $SerialPort -Address [convert]::ToUInt16($Address, 16)
            }
            3 {
                $BaudRate = Read-Host "Enter new baud rate (0-9)"
                Change-Baud-Rate -SerialPort $SerialPort -BaudRate [byte]::Parse($BaudRate)
            }
            4 {
                $AutoUpload = Read-Host "Auto upload (true/false)"
                Set-Upload-Mode -SerialPort $SerialPort -AutoUpload [bool]::Parse($AutoUpload)
            }
            5 {
                $Interval = Read-Host "Enter upload interval (1-100)"
                Set-Upload-Interval -SerialPort $SerialPort -Interval [byte]::Parse($Interval)
            }
            6 {
                $Mode = Read-Host "Enter LED mode (0-3)"
                Set-LED-Mode -SerialPort $SerialPort -Mode [byte]::Parse($Mode)
            }
            7 {
                $Mode = Read-Host "Enter relay mode (0-1)"
                Set-Relay-Mode -SerialPort $SerialPort -Mode [byte]::Parse($Mode)
            }
            8 {
                $Mode = Read-Host "Enter communication mode (0=UART, 1=Relay)"
                Set-Communication-Mode -SerialPort $SerialPort -Mode [byte]::Parse($Mode)
            }
            9 {
                $Distance = Read-Distance -SerialPort $SerialPort
                if ($Distance -ne $null) {
                    Write-Host "Distance: $Distance mm"
                }
            }
            10 {
                $SerialPort.Close()
                $COMPort = Read-Host "Enter new COM port"
                $SerialPort.PortName = $COMPort
                $SerialPort.Open()
            }
            11 {
                $SerialPort.Close()
                return
            }
            default {
                Write-Host "Invalid option"
            }
        }

        Write-Host "Press Enter to continue..."
        [void][System.Console]::ReadLine()
    }
}

$COMPort = Read-Host "Enter COM port"
$BaudRate = Read-Host "Enter baud rate"

$SerialPort = New-Object System.IO.Ports.SerialPort $COMPort, $BaudRate, [System.IO.Ports.Parity]::None, 8, [System.IO.Ports.StopBits]::One
Initialize-Sensor -SerialPort $SerialPort -BaudRate $BaudRate

Show-Menu -SerialPort $SerialPort
