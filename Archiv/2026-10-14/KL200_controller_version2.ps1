<#
.SYNOPSIS
    XKC_KL200 Sensor Management Tool
.DESCRIPTION
    PowerShell tool for managing XKC_KL200 ultrasonic distance sensors
.NOTES
    Author: AI Assistant
    Date: 2025-04-21
    Version: 1.0
#>

# Constants for error codes
$XKC_SUCCESS = 0
$XKC_INVALID_PARAMETER = 1
$XKC_TIMEOUT = 2
$XKC_CHECKSUM_ERROR = 3
$XKC_RESPONSE_ERROR = 4

# Global variables
$script:serialPort = $null
$script:autoMode = $false
$script:lastReceivedDistance = 0
$script:distance = 0
$script:available = $false

function Initialize-SerialPort {
    param (
        [string]$PortName,
        [int]$BaudRate = 9600
    )

    try {
        # Load System.IO.Ports assembly if not already loaded
        if (-not ([System.Management.Automation.PSTypeName]'System.IO.Ports.SerialPort').Type) {
            Add-Type -AssemblyName System.IO.Ports
        }

        # Create and configure serial port
        $script:serialPort = New-Object System.IO.Ports.SerialPort
        $script:serialPort.PortName = $PortName
        $script:serialPort.BaudRate = $BaudRate
        $script:serialPort.DataBits = 8
        $script:serialPort.Parity = [System.IO.Ports.Parity]::None
        $script:serialPort.StopBits = [System.IO.Ports.StopBits]::One
        $script:serialPort.ReadTimeout = 1000
        $script:serialPort.WriteTimeout = 1000
        
        # Open the port
        $script:serialPort.Open()
        
        Write-Host "Serial port $PortName initialized at $BaudRate baud" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error initializing serial port: $_" -ForegroundColor Red
        return $false
    }
}

function Close-SerialPort {
    if ($script:serialPort -and $script:serialPort.IsOpen) {
        $script:serialPort.Close()
        $script:serialPort.Dispose()
        $script:serialPort = $null
        Write-Host "Serial port closed" -ForegroundColor Green
    }
}

function Send-Command {
    param (
        [byte[]]$Command
    )
    
    if (-not $script:serialPort -or -not $script:serialPort.IsOpen) {
        Write-Host "Serial port not open" -ForegroundColor Red
        return $false
    }
    
    try {
        $script:serialPort.Write($Command, 0, $Command.Length)
        return $true
    }
    catch {
        Write-Host "Error sending command: $_" -ForegroundColor Red
        return $false
    }
}

function Calculate-Checksum {
    param (
        [byte[]]$Data,
        [int]$Length
    )
    
    $checksum = 0
    for ($i = 0; $i -lt $Length; $i++) {
        $checksum = $checksum -bxor $Data[$i]
    }
    
    return $checksum
}

function Wait-ForResponse {
    param (
        [byte]$ExpectedCmd,
        [int]$Timeout = 1000
    )
    
    if (-not $script:serialPort -or -not $script:serialPort.IsOpen) {
        Write-Host "Serial port not open" -ForegroundColor Red
        return $XKC_RESPONSE_ERROR
    }
    
    try {
        $startTime = Get-Date
        $response = New-Object byte[] 9
        
        # Wait for data with timeout
        while ($script:serialPort.BytesToRead -lt 9) {
            if (((Get-Date) - $startTime).TotalMilliseconds -gt $Timeout) {
                Write-Host "Timeout waiting for response" -ForegroundColor Yellow
                return $XKC_TIMEOUT
            }
            Start-Sleep -Milliseconds 10
        }
        
        # Read response
        $script:serialPort.Read($response, 0, 9)
        
        # Check if response matches expected command
        if ($response[0] -eq 0x62 -and $response[1] -eq $ExpectedCmd) {
            $checksum = $response[8]
            $calcChecksum = Calculate-Checksum -Data $response -Length 8
            
            if ($checksum -eq $calcChecksum) {
                # Check if response indicates successful execution (0x66)
                if ($response[7] -eq 0x66) {
                    return $XKC_SUCCESS
                }
            }
            else {
                return $XKC_CHECKSUM_ERROR
            }
        }
        
        return $XKC_RESPONSE_ERROR
    }
    catch {
        Write-Host "Error waiting for response: $_" -ForegroundColor Red
        return $XKC_RESPONSE_ERROR
    }
}

function Invoke-HardReset {
    Write-Host "Performing hard reset (factory reset)..." -ForegroundColor Cyan
    
    # Command to reset to factory settings
    $command = [byte[]]@(0x62, 0x39, 0x09, 0xFF, 0xFF, 0xFF, 0xFF, 0xFE, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
    if (Send-Command -Command $command) {
        $result = Wait-ForResponse -ExpectedCmd 0x39
        
        switch ($result) {
            $XKC_SUCCESS { Write-Host "Hard reset successful" -ForegroundColor Green; return $true }
            $XKC_TIMEOUT { Write-Host "Timeout during hard reset" -ForegroundColor Yellow }
            $XKC_CHECKSUM_ERROR { Write-Host "Checksum error during hard reset" -ForegroundColor Red }
            $XKC_RESPONSE_ERROR { Write-Host "Invalid response during hard reset" -ForegroundColor Red }
            default { Write-Host "Unknown error during hard reset" -ForegroundColor Red }
        }
    }
    
    return $false
}

function Invoke-SoftReset {
    Write-Host "Performing soft reset (user settings reset)..." -ForegroundColor Cyan
    
    # Command to reset to user settings
    $command = [byte[]]@(0x62, 0x39, 0x09, 0xFF, 0xFF, 0xFF, 0xFF, 0xFD, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
    if (Send-Command -Command $command) {
        $result = Wait-ForResponse -ExpectedCmd 0x39
        
        switch ($result) {
            $XKC_SUCCESS { Write-Host "Soft reset successful" -ForegroundColor Green; return $true }
            $XKC_TIMEOUT { Write-Host "Timeout during soft reset" -ForegroundColor Yellow }
            $XKC_CHECKSUM_ERROR { Write-Host "Checksum error during soft reset" -ForegroundColor Red }
            $XKC_RESPONSE_ERROR { Write-Host "Invalid response during soft reset" -ForegroundColor Red }
            default { Write-Host "Unknown error during soft reset" -ForegroundColor Red }
        }
    }
    
    return $false
}

function Set-SensorAddress {
    param (
        [uint16]$Address
    )
    
    if ($Address -gt 0xFFFE) {
        Write-Host "Invalid address. Must be between 0 and 65534" -ForegroundColor Red
        return $false
    }
    
    Write-Host "Changing sensor address to $Address..." -ForegroundColor Cyan
    
    # Command to change address
    $command = [byte[]]@(0x62, 0x32, 0x09, 0xFF, 0xFF, [byte]($Address -shr 8), [byte]$Address, 0x00, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
    if (Send-Command -Command $command) {
        $result = Wait-ForResponse -ExpectedCmd 0x32
        
        switch ($result) {
            $XKC_SUCCESS { Write-Host "Address change successful" -ForegroundColor Green; return $true }
            $XKC_TIMEOUT { Write-Host "Timeout during address change" -ForegroundColor Yellow }
            $XKC_CHECKSUM_ERROR { Write-Host "Checksum error during address change" -ForegroundColor Red }
            $XKC_RESPONSE_ERROR { Write-Host "Invalid response during address change" -ForegroundColor Red }
            default { Write-Host "Unknown error during address change" -ForegroundColor Red }
        }
    }
    
    return $false
}

function Set-BaudRate {
    param (
        [byte]$BaudRateIndex
    )
    
    if ($BaudRateIndex -gt 9) {
        Write-Host "Invalid baud rate index. Must be between 0 and 9" -ForegroundColor Red
        return $false
    }
    
    $baudRates = @(1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200, 230400, 460800)
    Write-Host "Changing baud rate to $($baudRates[$BaudRateIndex])..." -ForegroundColor Cyan
    
    # Command to change baud rate
    $command = [byte[]]@(0x62, 0x30, 0x09, 0xFF, 0xFF, 0x00, $BaudRateIndex, 0x00, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
    if (Send-Command -Command $command) {
        $result = Wait-ForResponse -ExpectedCmd 0x30
        
        switch ($result) {
            $XKC_SUCCESS { 
                Write-Host "Baud rate change successful" -ForegroundColor Green
                Write-Host "Reconnecting with new baud rate..." -ForegroundColor Cyan
                
                # Close and reopen the port with the new baud rate
                $portName = $script:serialPort.PortName
                Close-SerialPort
                Start-Sleep -Milliseconds 500
                Initialize-SerialPort -PortName $portName -BaudRate $baudRates[$BaudRateIndex]
                
                return $true 
            }
            $XKC_TIMEOUT { Write-Host "Timeout during baud rate change" -ForegroundColor Yellow }
            $XKC_CHECKSUM_ERROR { Write-Host "Checksum error during baud rate change" -ForegroundColor Red }
            $XKC_RESPONSE_ERROR { Write-Host "Invalid response during baud rate change" -ForegroundColor Red }
            default { Write-Host "Unknown error during baud rate change" -ForegroundColor Red }
        }
    }
    
    return $false
}

function Set-UploadMode {
    param (
        [bool]$AutoUpload
    )
    
    $mode = [byte]($AutoUpload -eq $true ? 1 : 0)
    $modeText = $AutoUpload ? "automatic" : "manual"
    
    Write-Host "Setting upload mode to $modeText..." -ForegroundColor Cyan
    
    # Command to set upload mode
    $command = [byte[]]@(0x62, 0x34, 0x09, 0xFF, 0xFF, 0x00, $mode, 0x00, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
    if (Send-Command -Command $command) {
        $result = Wait-ForResponse -ExpectedCmd 0x34
        
        switch ($result) {
            $XKC_SUCCESS { 
                Write-Host "Upload mode change successful" -ForegroundColor Green
                $script:autoMode = $AutoUpload
                return $true 
            }
            $XKC_TIMEOUT { Write-Host "Timeout during upload mode change" -ForegroundColor Yellow }
            $XKC_CHECKSUM_ERROR { Write-Host "Checksum error during upload mode change" -ForegroundColor Red }
            $XKC_RESPONSE_ERROR { Write-Host "Invalid response during upload mode change" -ForegroundColor Red }
            default { Write-Host "Unknown error during upload mode change" -ForegroundColor Red }
        }
    }
    
    return $false
}

function Set-UploadInterval {
    param (
        [byte]$Interval
    )
    
    if ($Interval -lt 1 -or $Interval -gt 100) {
        Write-Host "Invalid interval. Must be between 1 and 100" -ForegroundColor Red
        return $false
    }
    
    Write-Host "Setting upload interval to $Interval (${Interval}00ms)..." -ForegroundColor Cyan
    
    # Command to set upload interval
    $command = [byte[]]@(0x62, 0x35, 0x09, 0xFF, 0xFF, 0x00, $Interval, 0x00, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
    if (Send-Command -Command $command) {
        $result = Wait-ForResponse -ExpectedCmd 0x35
        
        switch ($result) {
            $XKC_SUCCESS { Write-Host "Upload interval change successful" -ForegroundColor Green; return $true }
            $XKC_TIMEOUT { Write-Host "Timeout during upload interval change" -ForegroundColor Yellow }
            $XKC_CHECKSUM_ERROR { Write-Host "Checksum error during upload interval change" -ForegroundColor Red }
            $XKC_RESPONSE_ERROR { Write-Host "Invalid response during upload interval change" -ForegroundColor Red }
            default { Write-Host "Unknown error during upload interval change" -ForegroundColor Red }
        }
    }
    
    return $false
}

function Set-LEDMode {
    param (
        [byte]$Mode
    )
    
    if ($Mode -gt 3) {
        Write-Host "Invalid LED mode. Must be between 0 and 3" -ForegroundColor Red
        return $false
    }
    
    $modeText = switch ($Mode) {
        0 { "on when detecting" }
        1 { "off when detecting" }
        2 { "always off" }
        3 { "always on" }
    }
    
    Write-Host "Setting LED mode to $Mode ($modeText)..." -ForegroundColor Cyan
    
    # Command to set LED mode
    $command = [byte[]]@(0x62, 0x37, 0x09, 0xFF, 0xFF, 0x00, $Mode, 0x00, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
     if (Send-Command -Command $command) {
        $result = Wait-ForResponse -ExpectedCmd 0x37
        
        switch ($result) {
            $XKC_SUCCESS { Write-Host "LED mode change successful" -ForegroundColor Green; return $true }
            $XKC_TIMEOUT { Write-Host "Timeout during LED mode change" -ForegroundColor Yellow }
            $XKC_CHECKSUM_ERROR { Write-Host "Checksum error during LED mode change" -ForegroundColor Red }
            $XKC_RESPONSE_ERROR { Write-Host "Invalid response during LED mode change" -ForegroundColor Red }
            default { Write-Host "Unknown error during LED mode change" -ForegroundColor Red }
        }
    }
    
    return $false
}

function Set-RelayMode {
    param (
        [byte]$Mode
    )
    
    if ($Mode -gt 1) {
        Write-Host "Invalid relay mode. Must be 0 or 1" -ForegroundColor Red
        return $false
    }
    
    $modeText = $Mode -eq 0 ? "active when detecting" : "inactive when detecting"
    Write-Host "Setting relay mode to $Mode ($modeText)..." -ForegroundColor Cyan
    
    # Command to set relay mode
    $command = [byte[]]@(0x62, 0x38, 0x09, 0xFF, 0xFF, 0x00, $Mode, 0x00, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
    if (Send-Command -Command $command) {
        $result = Wait-ForResponse -ExpectedCmd 0x38
        
        switch ($result) {
            $XKC_SUCCESS { Write-Host "Relay mode change successful" -ForegroundColor Green; return $true }
            $XKC_TIMEOUT { Write-Host "Timeout during relay mode change" -ForegroundColor Yellow }
            $XKC_CHECKSUM_ERROR { Write-Host "Checksum error during relay mode change" -ForegroundColor Red }
            $XKC_RESPONSE_ERROR { Write-Host "Invalid response during relay mode change" -ForegroundColor Red }
            default { Write-Host "Unknown error during relay mode change" -ForegroundColor Red }
        }
    }
    
    return $false
}

function Set-CommunicationMode {
    param (
        [byte]$Mode
    )
    
    if ($Mode -gt 1) {
        Write-Host "Invalid communication mode. Must be 0 or 1" -ForegroundColor Red
        return $false
    }
    
    $modeText = $Mode -eq 0 ? "relay mode" : "UART mode"
    Write-Host "Setting communication mode to $Mode ($modeText)..." -ForegroundColor Cyan
    
    # Command to set communication mode
    $command = [byte[]]@(0x61, 0x30, 0x09, 0xFF, 0xFF, 0x00, $Mode, 0x00, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
    if (Send-Command -Command $command) {
        $result = Wait-ForResponse -ExpectedCmd 0x30
        
        switch ($result) {
            $XKC_SUCCESS { Write-Host "Communication mode change successful" -ForegroundColor Green; return $true }
            $XKC_TIMEOUT { Write-Host "Timeout during communication mode change" -ForegroundColor Yellow }
            $XKC_CHECKSUM_ERROR { Write-Host "Checksum error during communication mode change" -ForegroundColor Red }
            $XKC_RESPONSE_ERROR { Write-Host "Invalid response during communication mode change" -ForegroundColor Red }
            default { Write-Host "Unknown error during communication mode change" -ForegroundColor Red }
        }
    }
    
    return $false
}

function Read-Distance {
    param (
        [int]$Timeout = 1000
    )
    
    if ($script:autoMode) {
        Write-Host "In auto mode, use Process-AutoData instead" -ForegroundColor Yellow
        return $script:lastReceivedDistance
    }
    
    Write-Host "Reading distance..." -ForegroundColor Cyan
    
    # Command to read distance
    $command = [byte[]]@(0x62, 0x33, 0x09, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00)
    $command[8] = Calculate-Checksum -Data $command -Length 8
    
    if (-not (Send-Command -Command $command)) {
        return $script:lastReceivedDistance
    }
    
    try {
        $startTime = Get-Date
        $response = New-Object byte[] 9
        
        # Wait for data with timeout
        while ($script:serialPort.BytesToRead -lt 9) {
            if (((Get-Date) - $startTime).TotalMilliseconds -gt $Timeout) {
                Write-Host "Timeout waiting for distance data" -ForegroundColor Yellow
                return $script:lastReceivedDistance
            }
            Start-Sleep -Milliseconds 10
        }
        
        # Read response
        $script:serialPort.Read($response, 0, 9)
        
        if ($response[0] -eq 0x62 -and $response[1] -eq 0x33) {
            $length = $response[2]
            $address = ($response[3] -shl 8) -bor $response[4]
            $rawDistance = ($response[5] -shl 8) -bor $response[6]
            $checksum = $response[8]
            $calcChecksum = Calculate-Checksum -Data $response -Length 8
            
            if ($checksum -eq $calcChecksum) {
                $script:distance = $rawDistance
                $script:lastReceivedDistance = $rawDistance
                $script:available = $true
                Write-Host "Distance: $rawDistance mm" -ForegroundColor Green
                return $script:distance
            }
            else {
                Write-Host "Checksum error in distance data" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Invalid response format for distance data" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error reading distance: $_" -ForegroundColor Red
    }
    
    return $script:lastReceivedDistance
}

function Process-AutoData {
    if (-not $script:autoMode -or $script:serialPort.BytesToRead -lt 9) {
        return $false
    }
    
    try {
        $response = New-Object byte[] 9
        $script:serialPort.Read($response, 0, 9)
        
        if ($response[0] -eq 0x62 -and $response[1] -eq 0x33) {
            $length = $response[2]
            $address = ($response[3] -shl 8) -bor $response[4]
            $rawDistance = ($response[5] -shl 8) -bor $response[6]
            $checksum = $response[8]
            $calcChecksum = Calculate-Checksum -Data $response -Length 8
            
            if ($checksum -eq $calcChecksum) {
                $script:distance = $rawDistance
                $script:lastReceivedDistance = $rawDistance
                $script:available = $true
                Write-Host "Auto distance: $rawDistance mm" -ForegroundColor Green
                return $true
            }
        }
        
        # Invalid data, discard one byte
        if ($script:serialPort.BytesToRead -gt 0) {
            $script:serialPort.ReadByte() | Out-Null
        }
    }
    catch {
        Write-Host "Error processing auto data: $_" -ForegroundColor Red
    }
    
    return $false
}

function Monitor-Distance {
    param (
        [int]$Duration = 30,  # Duration in seconds
        [int]$Interval = 500  # Interval in milliseconds
    )
    
    Write-Host "Starting distance monitoring for $Duration seconds..." -ForegroundColor Cyan
    
    $startTime = Get-Date
    $endTime = $startTime.AddSeconds($Duration)
    
    while ((Get-Date) -lt $endTime) {
        if ($script:autoMode) {
            # In auto mode, just process any available data
            while ($script:serialPort.BytesToRead -ge 9) {
                Process-AutoData
            }
        }
        else {
            # In manual mode, actively request distance
            Read-Distance
        }
        
        Start-Sleep -Milliseconds $Interval
        
        # Show countdown
        $remainingTime = ($endTime - (Get-Date)).TotalSeconds
        Write-Host "`rMonitoring: $([Math]::Round($remainingTime)) seconds remaining..." -NoNewline -ForegroundColor Cyan
    }
    
    Write-Host "`nMonitoring complete" -ForegroundColor Green
}

function Show-Menu {
    Clear-Host
    Write-Host "===== XKC_KL200 Sensor Management Tool =====" -ForegroundColor Cyan
    Write-Host
    Write-Host "Connection Status: " -NoNewline
    
    if ($script:serialPort -and $script:serialPort.IsOpen) {
        Write-Host "Connected to $($script:serialPort.PortName) at $($script:serialPort.BaudRate) baud" -ForegroundColor Green
    }
    else {
        Write-Host "Not connected" -ForegroundColor Red
    }
    
    Write-Host "Upload Mode: " -NoNewline
    if ($script:autoMode) {
        Write-Host "Automatic" -ForegroundColor Yellow
    }
    else {
        Write-Host "Manual" -ForegroundColor Yellow
    }
    
    Write-Host "Last Distance: $script:lastReceivedDistance mm" -ForegroundColor Yellow
    Write-Host
    Write-Host "1. Connect to Sensor" -ForegroundColor White
    Write-Host "2. Disconnect" -ForegroundColor White
    Write-Host "3. Read Distance (Manual Mode)" -ForegroundColor White
    Write-Host "4. Monitor Distance" -ForegroundColor White
    Write-Host "5. Hard Reset (Factory Reset)" -ForegroundColor White
    Write-Host "6. Soft Reset (User Settings Reset)" -ForegroundColor White
    Write-Host "7. Change Sensor Address" -ForegroundColor White
    Write-Host "8. Change Baud Rate" -ForegroundColor White
    Write-Host "9. Set Upload Mode" -ForegroundColor White
    Write-Host "10. Set Upload Interval" -ForegroundColor White
    Write-Host "11. Set LED Mode" -ForegroundColor White
    Write-Host "12. Set Relay Mode" -ForegroundColor White
    Write-Host "13. Set Communication Mode" -ForegroundColor White
    Write-Host "Q. Quit" -ForegroundColor White
    Write-Host
    Write-Host "Enter your choice: " -NoNewline -ForegroundColor Cyan
}

function Main {
    try {
        do {
            Show-Menu
            $choice = Read-Host
            
            switch ($choice) {
                "1" {
                    # Connect to sensor
                    $availablePorts = [System.IO.Ports.SerialPort]::GetPortNames()
                    
                    if ($availablePorts.Count -eq 0) {
                        Write-Host "No serial ports found" -ForegroundColor Red
                        break
                    }
                    
                    Write-Host "Available ports:" -ForegroundColor Cyan
                    for ($i = 0; $i -lt $availablePorts.Count; $i++) {
                        Write-Host "$($i+1). $($availablePorts[$i])" -ForegroundColor White
                    }
                    
                    $portIndex = Read-Host "Select port number"
                    if ([int]::TryParse($portIndex, [ref]$null)) {
                        $portIndex = [int]$portIndex - 1
                        if ($portIndex -ge 0 -and $portIndex -lt $availablePorts.Count) {
                            $portName = $availablePorts[$portIndex]
                            
                            $baudRates = @(1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200, 230400, 460800)
                            Write-Host "Available baud rates:" -ForegroundColor Cyan
                            for ($i = 0; $i -lt $baudRates.Count; $i++) {
                                Write-Host "$($i+1). $($baudRates[$i])" -ForegroundColor White
                            }
                            
                            $baudIndex = Read-Host "Select baud rate number (default: 4 - 9600)"
                            if ([string]::IsNullOrEmpty($baudIndex)) {
                                $baudRate = 9600
                            }
                            elseif ([int]::TryParse($baudIndex, [ref]$null)) {
                                $baudIndex = [int]$baudIndex - 1
                                if ($baudIndex -ge 0 -and $baudIndex -lt $baudRates.Count) {
                                    $baudRate = $baudRates[$baudIndex]
                                }
                                else {
                                    $baudRate = 9600
                                }
                            }
                            else {
                                $baudRate = 9600
                            }
                            
                            # Close existing connection if any
                            Close-SerialPort
                            
                            # Initialize new connection
                            Initialize-SerialPort -PortName $portName -BaudRate $baudRate
                        }
                        else {
                            Write-Host "Invalid port selection" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "Invalid input" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "2" {
                    # Disconnect
                    Close-SerialPort
                    Read-Host "Press Enter to continue"
                }
                
                "3" {
                    # Read distance
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        Read-Distance
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                "4" {
                    # Monitor distance
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        $duration = Read-Host "Enter monitoring duration in seconds (default: 30)"
                        if ([string]::IsNullOrEmpty($duration)) {
                            $duration = 30
                        }
                        elseif (-not [int]::TryParse($duration, [ref]$null)) {
                            $duration = 30
                        }
                        
                        $interval = Read-Host "Enter reading interval in milliseconds (default: 500)"
                        if ([string]::IsNullOrEmpty($interval)) {
                            $interval = 500
                        }
                        elseif (-not [int]::TryParse($interval, [ref]$null)) {
                            $interval = 500
                        }
                        
                        Monitor-Distance -Duration $duration -Interval $interval
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "5" {
                    # Hard reset
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        $confirm = Read-Host "Are you sure you want to perform a factory reset? (y/n)"
                        if ($confirm -eq "y" -or $confirm -eq "Y") {
                            Invoke-HardReset
                        }
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "6" {
                    # Soft reset
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        $confirm = Read-Host "Are you sure you want to perform a user settings reset? (y/n)"
                        if ($confirm -eq "y" -or $confirm -eq "Y") {
                            Invoke-SoftReset
                        }
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "7" {
                    # Change sensor address
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        $address = Read-Host "Enter new address (0-65534)"
                        if ([uint16]::TryParse($address, [ref]$null)) {
                            $addressValue = [uint16]$address
                            if ($addressValue -le 0xFFFE) {
                                Set-SensorAddress -Address $addressValue
                            }
                            else {
                                Write-Host "Invalid address. Must be between 0 and 65534" -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "Invalid input" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "8" {
                    # Change baud rate
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        $baudRates = @(1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200, 230400, 460800)
                        Write-Host "Available baud rates:" -ForegroundColor Cyan
                        for ($i = 0; $i -lt $baudRates.Count; $i++) {
                            Write-Host "$($i). $($baudRates[$i])" -ForegroundColor White
                        }
                        
                        $baudIndex = Read-Host "Select baud rate index (0-9)"
                        if ([byte]::TryParse($baudIndex, [ref]$null)) {
                            $baudIndexValue = [byte]$baudIndex
                            if ($baudIndexValue -le 9) {
                                Set-BaudRate -BaudRateIndex $baudIndexValue
                            }
                            else {
                                Write-Host "Invalid baud rate index. Must be between 0 and 9" -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "Invalid input" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "9" {
                    # Set upload mode
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        Write-Host "Upload modes:" -ForegroundColor Cyan
                        Write-Host "0. Manual (query mode)" -ForegroundColor White
                        Write-Host "1. Automatic (continuous upload)" -ForegroundColor White
                        
                        $mode = Read-Host "Select upload mode (0-1)"
                        if ($mode -eq "0") {
                            Set-UploadMode -AutoUpload $false
                        }
                        elseif ($mode -eq "1") {
                            Set-UploadMode -AutoUpload $true
                        }
                        else {
                            Write-Host "Invalid input" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "10" {
                    # Set upload interval
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        $interval = Read-Host "Enter upload interval (1-100, corresponds to 100ms-10s)"
                        if ([byte]::TryParse($interval, [ref]$null)) {
                            $intervalValue = [byte]$interval
                            if ($intervalValue -ge 1 -and $intervalValue -le 100) {
                                Set-UploadInterval -Interval $intervalValue
                            }
                            else {
                                Write-Host "Invalid interval. Must be between 1 and 100" -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "Invalid input" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "11" {
                    # Set LED mode
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        Write-Host "LED modes:" -ForegroundColor Cyan
                        Write-Host "0. On when detecting" -ForegroundColor White
                        Write-Host "1. Off when detecting" -ForegroundColor White
                        Write-Host "2. Always off" -ForegroundColor White
                        Write-Host "3. Always on" -ForegroundColor White
                        
                        $mode = Read-Host "Select LED mode (0-3)"
                        if ([byte]::TryParse($mode, [ref]$null)) {
                            $modeValue = [byte]$mode
                            if ($modeValue -le 3) {
                                Set-LEDMode -Mode $modeValue
                            }
                            else {
                                Write-Host "Invalid mode. Must be between 0 and 3" -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "Invalid input" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "12" {
                    # Set relay mode
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        Write-Host "Relay modes:" -ForegroundColor Cyan
                        Write-Host "0. Active when detecting" -ForegroundColor White
                        Write-Host "1. Inactive when detecting" -ForegroundColor White
                        
                        $mode = Read-Host "Select relay mode (0-1)"
                        if ([byte]::TryParse($mode, [ref]$null)) {
                            $modeValue = [byte]$mode
                            if ($modeValue -le 1) {
                                Set-RelayMode -Mode $modeValue
                            }
                            else {
                                Write-Host "Invalid mode. Must be 0 or 1" -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "Invalid input" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "13" {
                    # Set communication mode
                    if ($script:serialPort -and $script:serialPort.IsOpen) {
                        Write-Host "Communication modes:" -ForegroundColor Cyan
                        Write-Host "0. Relay mode" -ForegroundColor White
                        Write-Host "1. UART mode" -ForegroundColor White
                        
                        $mode = Read-Host "Select communication mode (0-1)"
                        if ([byte]::TryParse($mode, [ref]$null)) {
                            $modeValue = [byte]$mode
                            if ($modeValue -le 1) {
                                Set-CommunicationMode -Mode $modeValue
                            }
                            else {
                                Write-Host "Invalid mode. Must be 0 or 1" -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "Invalid input" -ForegroundColor Red
                        }
                    }
                    else {
                        Write-Host "Not connected to sensor" -ForegroundColor Red
                    }
                    
                    Read-Host "Press Enter to continue"
                }
                
                "q" { 
                    # Quit
                    Write-Host "Exiting..." -ForegroundColor Cyan 
                }
                
                "Q" { 
                    # Quit
                    Write-Host "Exiting..." -ForegroundColor Cyan 
                }
                
                default { 
                    Write-Host "Invalid choice. Please try again." -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                }
            }
        } while ($choice -ne "q" -and $choice -ne "Q")
    }
    finally {
        # Ensure serial port is closed when script exits
        Close-SerialPort
    }
}

# Start the main program
Main
