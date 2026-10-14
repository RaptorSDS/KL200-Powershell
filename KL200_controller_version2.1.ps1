#Requires -Version 5.1
<#
.SYNOPSIS
    XKC-KL200 Laser Sensor Management Tool (Version 2.1)
.DESCRIPTION
    PowerShell tool for configuring and reading XKC-KL200-2M-UART laser
    distance sensors (time of flight, 10-2000 mm) via a serial port.
    Runs on Windows PowerShell 5.1 and PowerShell 7+.

    Protocol (datasheet V10, 2023-03-28): 9 byte frames
      [0x62|0x61] [Cmd] [Len] [AddrHi] [AddrLo] [DataHi] [DataLo] [Answer] [XOR]
    Successful command answer: 0x62 [Cmd] 0x09 00 00 00 00 0x66 [XOR]
.PARAMETER Port
    Serial port to connect to at start (e.g. COM3). Without it the menu starts unconnected.
.PARAMETER BaudRate
    Baud rate used together with -Port (default 9600 = sensor factory default).
.PARAMETER Address
    Sensor address used in commands (default 0xFFFF = broadcast / any address).
.PARAMETER Action
    Menu (default), ReadDistance, ReadConfig or Monitor. All except Menu require -Port.
.PARAMETER Duration
    Monitor duration in seconds (Action Monitor).
.PARAMETER Interval
    Monitor interval in milliseconds in manual upload mode (Action Monitor).
.PARAMETER CsvPath
    Optional CSV file for monitor readings (Action Monitor).
.EXAMPLE
    .\KL200_controller_version2.1.ps1
.EXAMPLE
    .\KL200_controller_version2.1.ps1 -Port COM3 -Action ReadDistance
.EXAMPLE
    .\KL200_controller_version2.1.ps1 -Port COM3 -Action Monitor -Duration 60 -CsvPath .\distance.csv
.NOTES
    Version: 2.1
    Date:    2026-10-14
#>
[CmdletBinding()]
param (
    [string]$Port,

    [ValidateSet(2400, 4800, 9600, 14400, 19200, 38400, 56000, 57600, 115200, 128000)]
    [int]$BaudRate = 9600,

    [ValidateRange(0, 65535)]
    [int]$Address = 0xFFFF,

    [ValidateSet('Menu', 'ReadDistance', 'ReadConfig', 'Monitor')]
    [string]$Action = 'Menu',

    [ValidateRange(1, 86400)]
    [int]$Duration = 30,

    [ValidateRange(50, 60000)]
    [int]$Interval = 500,

    [string]$CsvPath
)

# ---------------------------------------------------------------------------
# Constants (datasheet XKC-KL200-2M-UART)
# ---------------------------------------------------------------------------

enum XkcStatus {
    Success
    Timeout
    ChecksumError
    ResponseError
    NotConnected
}

# Index = value sent to the sensor (command 0x30)
$script:BaudRates = @(2400, 4800, 9600, 14400, 19200, 38400, 56000, 57600, 115200, 128000)
$script:DefaultBaudIndex = 2

$script:FrameHeader     = [byte]0x62
$script:LineModeHeader  = [byte]0x61   # line mode / response speed are sent with 0x61
$script:AnswerOk        = [byte]0x66

$script:Cmd = @{
    BaudRate       = [byte]0x30   # b0  (with header 0x61: line mode a0)
    ReadConfig     = [byte]0x31   # b1
    Address        = [byte]0x32   # b2  (with header 0x61: response speed a2)
    Distance       = [byte]0x33   # b3
    UploadMode     = [byte]0x34   # b4
    UploadInterval = [byte]0x35   # b5
    LedMode        = [byte]0x37   # b7
    RelayMode      = [byte]0x38   # b8
    Reset          = [byte]0x39   # b9
}

$script:UploadModes   = @('Manual (query)', 'Automatic (continuous upload)')
$script:LedModes      = @('On when object detected', 'Off when object detected', 'Always off', 'Always on')
$script:RelayModes    = @('Active when object detected', 'Inactive when object detected')
$script:LineModes     = @('Relay mode (switch output, no UART data)', 'UART mode')
$script:ResponseModes = @('High precision', 'Normal', 'High stability')

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------

$script:SerialPort    = $null
$script:SensorAddress = [uint16]$Address
$script:LastDistance  = $null
$script:Config        = $null      # last configuration read from the sensor
$script:AutoMode      = $false
$script:UploadInterval = 10        # x 100 ms, used for read timeouts in auto mode

# ---------------------------------------------------------------------------
# Low level helpers
# ---------------------------------------------------------------------------

function Write-Status {
    param (
        [string]$Message,
        [ValidateSet('Info', 'Ok', 'Warn', 'Error', 'Plain')]
        [string]$Level = 'Plain'
    )
    $color = switch ($Level) {
        'Info'  { 'Cyan' }
        'Ok'    { 'Green' }
        'Warn'  { 'Yellow' }
        'Error' { 'Red' }
        default { 'White' }
    }
    Write-Host $Message -ForegroundColor $color
}

function Test-SensorConnection {
    param ([switch]$Quiet)
    $connected = ($null -ne $script:SerialPort) -and $script:SerialPort.IsOpen
    if (-not $connected -and -not $Quiet) {
        Write-Status 'Not connected to sensor' Error
    }
    return $connected
}

function Get-Checksum {
    param (
        [byte[]]$Data,
        [int]$Length = $Data.Length
    )
    $checksum = 0
    for ($i = 0; $i -lt $Length; $i++) {
        $checksum = $checksum -bxor $Data[$i]
    }
    return [byte]$checksum
}

function ConvertTo-HexString {
    param ([byte[]]$Data)
    return (($Data | ForEach-Object { '{0:X2}' -f $_ }) -join ' ')
}

function New-SensorFrame {
    param (
        [byte]$Header = $script:FrameHeader,
        [byte]$Command,
        [uint16]$Data = 0,
        [byte]$Answer = 0
    )
    $frame = [byte[]]@(
        $Header,
        $Command,
        0x09,
        [byte](($script:SensorAddress -shr 8) -band 0xFF),
        [byte]($script:SensorAddress -band 0xFF),
        [byte](($Data -shr 8) -band 0xFF),
        [byte]($Data -band 0xFF),
        $Answer,
        0x00
    )
    $frame[8] = Get-Checksum -Data $frame -Length 8
    return , $frame
}

function Send-SensorFrame {
    param (
        [byte[]]$Frame,
        [switch]$KeepInput
    )
    if (-not (Test-SensorConnection)) { return $false }
    try {
        # Drop stale bytes (e.g. auto upload frames) so the answer is not mixed up
        if (-not $KeepInput) { $script:SerialPort.DiscardInBuffer() }
        $script:SerialPort.Write($Frame, 0, $Frame.Length)
        Write-Verbose "TX: $(ConvertTo-HexString $Frame)"
        return $true
    }
    catch {
        Write-Status "Error sending command: $($_.Exception.Message)" Error
        return $false
    }
}

function Read-SensorFrame {
    <#
        Reads bytes until a valid frame "0x62 <Command> <Len> ... <XOR>" is found.
        Bytes before the header and frames for other commands are skipped, so the
        reader resynchronises on its own. Returns @{ Status; Frame }.
    #>
    param (
        [byte]$Command,
        [int]$TimeoutMs = 1000
    )

    $result = [pscustomobject]@{ Status = [XkcStatus]::Timeout; Frame = $null }
    if (-not (Test-SensorConnection -Quiet)) {
        $result.Status = [XkcStatus]::NotConnected
        return $result
    }

    $buffer = New-Object System.Collections.Generic.List[byte]
    $expectedLength = 0
    $sawBadChecksum = $false
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        while ($stopwatch.ElapsedMilliseconds -lt $TimeoutMs) {
            if ($script:SerialPort.BytesToRead -eq 0) {
                Start-Sleep -Milliseconds 5
                continue
            }
            $b = [byte]$script:SerialPort.ReadByte()

            # Note: no switch here, "continue" inside switch does not continue the loop
            if ($buffer.Count -eq 0) {
                if ($b -eq $script:FrameHeader) { $buffer.Add($b) }
                continue
            }
            if ($buffer.Count -eq 1) {
                if ($b -eq $Command) { $buffer.Add($b) }
                else {
                    # Other frame: restart, the byte itself may be a new header
                    $buffer.Clear()
                    if ($b -eq $script:FrameHeader) { $buffer.Add($b) }
                }
                continue
            }
            if ($buffer.Count -eq 2) {
                if ($b -lt 9 -or $b -gt 64) {
                    $buffer.Clear()
                    if ($b -eq $script:FrameHeader) { $buffer.Add($b) }
                    continue
                }
                $expectedLength = [int]$b
            }
            $buffer.Add($b)

            if ($buffer.Count -eq $expectedLength) {
                $frame = $buffer.ToArray()
                Write-Verbose "RX: $(ConvertTo-HexString $frame)"
                if ((Get-Checksum -Data $frame -Length ($expectedLength - 1)) -eq $frame[$expectedLength - 1]) {
                    $result.Status = [XkcStatus]::Success
                    $result.Frame = $frame
                    return $result
                }
                $sawBadChecksum = $true
                $buffer.Clear()
            }
        }
    }
    catch {
        Write-Status "Error reading from sensor: $($_.Exception.Message)" Error
        $result.Status = [XkcStatus]::ResponseError
        return $result
    }

    if ($sawBadChecksum) { $result.Status = [XkcStatus]::ChecksumError }
    return $result
}

function Invoke-SensorCommand {
    <#
        Sends a configuration command and checks the 0x66 answer.
        Returns $true on success, otherwise prints the reason and returns $false.
    #>
    param (
        [byte]$Header = $script:FrameHeader,
        [byte]$Command,
        [uint16]$Data = 0,
        [byte]$Answer = 0,
        [string]$Label,
        [int]$TimeoutMs = 1000
    )

    if (-not (Test-SensorConnection)) { return $false }

    $frame = New-SensorFrame -Header $Header -Command $Command -Data $Data -Answer $Answer
    if (-not (Send-SensorFrame -Frame $frame)) { return $false }

    # The answer always starts with 0x62, also for commands sent with 0x61
    $response = Read-SensorFrame -Command $Command -TimeoutMs $TimeoutMs
    $status = $response.Status
    if ($status -eq [XkcStatus]::Success -and $response.Frame[7] -ne $script:AnswerOk) {
        $status = [XkcStatus]::ResponseError
    }

    switch ($status) {
        ([XkcStatus]::Success)       { Write-Status "$Label successful" Ok; return $true }
        ([XkcStatus]::Timeout)       { Write-Status "Timeout during $Label" Warn }
        ([XkcStatus]::ChecksumError) { Write-Status "Checksum error during $Label" Error }
        ([XkcStatus]::ResponseError) { Write-Status "Invalid response during $Label" Error }
        default                      { Write-Status "Error during $Label ($status)" Error }
    }
    return $false
}

# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------

function Connect-Sensor {
    param (
        [Parameter(Mandatory)]
        [string]$PortName,
        [int]$Baud = 9600,
        [switch]$SkipConfig
    )

    Disconnect-Sensor -Quiet

    try {
        $sp = New-Object System.IO.Ports.SerialPort
        $sp.PortName = $PortName
        $sp.BaudRate = $Baud
        $sp.DataBits = 8
        $sp.Parity = [System.IO.Ports.Parity]::None
        $sp.StopBits = [System.IO.Ports.StopBits]::One
        $sp.ReadTimeout = 1000
        $sp.WriteTimeout = 1000
        $sp.Open()
        $script:SerialPort = $sp
        Write-Status "Connected to $PortName at $Baud baud" Ok
    }
    catch {
        Write-Status "Error opening serial port: $($_.Exception.Message)" Error
        $script:SerialPort = $null
        return $false
    }

    if (-not $SkipConfig) { Update-SensorState }
    return $true
}

function Disconnect-Sensor {
    param ([switch]$Quiet)
    if ($null -ne $script:SerialPort) {
        try {
            if ($script:SerialPort.IsOpen) { $script:SerialPort.Close() }
            $script:SerialPort.Dispose()
        }
        catch { }
        $script:SerialPort = $null
        if (-not $Quiet) { Write-Status 'Serial port closed' Ok }
    }
}

function Update-SensorState {
    <#
        Reads the configuration to learn the real upload mode. If the sensor does not
        answer, detects auto mode by listening for unsolicited distance frames.
    #>
    $config = Get-SensorConfig -Quiet
    if ($null -ne $config) {
        $script:AutoMode = ($config.UploadMode -eq 1)
        $script:UploadInterval = [Math]::Max(1, [int]$config.UploadInterval)
        return
    }

    $script:SerialPort.DiscardInBuffer()
    $frame = Read-SensorFrame -Command $script:Cmd.Distance -TimeoutMs 1500
    $script:AutoMode = ($frame.Status -eq [XkcStatus]::Success)
    if ($script:AutoMode) {
        Write-Status 'Configuration not readable, but sensor sends data on its own -> automatic mode' Warn
    }
    else {
        Write-Status 'Configuration not readable, assuming manual mode' Warn
    }
}

# ---------------------------------------------------------------------------
# Sensor functions
# ---------------------------------------------------------------------------

function Get-SensorDistance {
    param (
        [int]$TimeoutMs = 0,
        [switch]$Quiet,
        [switch]$KeepInput     # auto mode monitoring: use frames already buffered
    )

    if (-not (Test-SensorConnection)) { return $null }

    if ($TimeoutMs -le 0) {
        $TimeoutMs = 1000
        if ($script:AutoMode) { $TimeoutMs = [Math]::Max(1000, $script:UploadInterval * 100 + 500) }
    }

    if ($script:AutoMode) {
        # Sensor uploads on its own; drop old frames to get a current value
        if (-not $KeepInput) { $script:SerialPort.DiscardInBuffer() }
    }
    else {
        $frame = New-SensorFrame -Command $script:Cmd.Distance
        if (-not (Send-SensorFrame -Frame $frame)) { return $null }
    }

    $response = Read-SensorFrame -Command $script:Cmd.Distance -TimeoutMs $TimeoutMs
    switch ($response.Status) {
        ([XkcStatus]::Success) {
            $distance = ([int]$response.Frame[5] -shl 8) -bor $response.Frame[6]
            $script:LastDistance = $distance
            if (-not $Quiet) { Write-Status "Distance: $distance mm" Ok }
            return $distance
        }
        ([XkcStatus]::Timeout)       { Write-Status 'Timeout waiting for distance data' Warn }
        ([XkcStatus]::ChecksumError) { Write-Status 'Checksum error in distance data' Error }
        default                      { Write-Status "Error reading distance ($($response.Status))" Error }
    }
    return $null
}

function Get-SensorConfig {
    param ([switch]$Quiet)

    if (-not (Test-SensorConnection)) { return $null }

    $frame = New-SensorFrame -Command $script:Cmd.ReadConfig
    if (-not (Send-SensorFrame -Frame $frame)) { return $null }

    $response = Read-SensorFrame -Command $script:Cmd.ReadConfig -TimeoutMs 1500
    if ($response.Status -ne [XkcStatus]::Success -or $response.Frame.Length -lt 29) {
        if (-not $Quiet) { Write-Status "Could not read configuration ($($response.Status))" Warn }
        return $null
    }

    $f = $response.Frame
    $config = [pscustomobject]@{
        BaudIndex           = [int]$f[11]
        BaudRate            = $(if ($f[11] -lt $script:BaudRates.Count) { $script:BaudRates[$f[11]] } else { $null })
        Address             = ([int]$f[12] -shl 8) -bor $f[13]
        CalibrationDistance = ([int]$f[14] -shl 8) -bor $f[15]
        UploadMode          = [int]$f[16]
        UploadInterval      = [int]$f[17]
        LedMode             = [int]$f[20]
        RelayMode           = [int]$f[21]
        LineMode            = [int]$f[22]
        Raw                 = ConvertTo-HexString $f
    }
    $script:Config = $config
    return $config
}

function Show-SensorConfig {
    $config = Get-SensorConfig
    if ($null -eq $config) { return }

    $script:AutoMode = ($config.UploadMode -eq 1)
    $script:UploadInterval = [Math]::Max(1, $config.UploadInterval)

    Write-Status 'Sensor configuration:' Info
    Write-Status ("  Baud rate:            {0} ({1})" -f $config.BaudIndex, $config.BaudRate)
    Write-Status ("  Device address:       {0} (0x{0:X4})" -f $config.Address)
    Write-Status ("  Calibration distance: {0} mm" -f $config.CalibrationDistance)
    Write-Status ("  Upload mode:          {0} ({1})" -f $config.UploadMode, (Get-ModeText $script:UploadModes $config.UploadMode))
    Write-Status ("  Upload interval:      {0} ({1} ms)" -f $config.UploadInterval, ($config.UploadInterval * 100))
    Write-Status ("  LED mode:             {0} ({1})" -f $config.LedMode, (Get-ModeText $script:LedModes $config.LedMode))
    Write-Status ("  Relay mode:           {0} ({1})" -f $config.RelayMode, (Get-ModeText $script:RelayModes $config.RelayMode))
    Write-Status ("  Line output mode:     {0} ({1})" -f $config.LineMode, (Get-ModeText $script:LineModes $config.LineMode))
    Write-Status ("  Raw:                  {0}" -f $config.Raw)
}

function Get-ModeText {
    param ([string[]]$Modes, [int]$Value)
    if ($Value -ge 0 -and $Value -lt $Modes.Count) { return $Modes[$Value] }
    return 'unknown'
}

function Reset-Sensor {
    param (
        [ValidateSet('Factory', 'User')]
        [string]$Mode
    )
    # FE = factory defaults (Xingkechuang), FD = user defaults; data is always 0xFFFF
    $answer = if ($Mode -eq 'Factory') { [byte]0xFE } else { [byte]0xFD }
    $label = if ($Mode -eq 'Factory') { 'factory reset' } else { 'user settings reset' }

    $ok = Invoke-SensorCommand -Command $script:Cmd.Reset -Data 0xFFFF -Answer $answer -Label $label
    if ($ok) {
        if ($Mode -eq 'Factory' -and $script:SerialPort.BaudRate -ne 9600) {
            Write-Status 'Factory default baud rate is 9600, reconnecting...' Info
            Start-Sleep -Milliseconds 500
            $null = Connect-Sensor -PortName $script:SerialPort.PortName -Baud 9600
        }
        else {
            Start-Sleep -Milliseconds 300
            Update-SensorState
        }
    }
    return $ok
}

function Set-SensorAddress {
    param (
        [ValidateRange(0, 65534)]
        [int]$NewAddress
    )
    $ok = Invoke-SensorCommand -Command $script:Cmd.Address -Data $NewAddress -Label "address change to $NewAddress"
    # Keep talking to the sensor: broadcast stays valid, a specific address must follow
    if ($ok -and $script:SensorAddress -ne 0xFFFF) { $script:SensorAddress = [uint16]$NewAddress }
    return $ok
}

function Set-SensorBaudRate {
    param (
        [ValidateRange(0, 9)]
        [int]$Index
    )
    $newBaud = $script:BaudRates[$Index]
    $ok = Invoke-SensorCommand -Command $script:Cmd.BaudRate -Data $Index -Label "baud rate change to $newBaud"
    if ($ok) {
        Write-Status 'Reconnecting with new baud rate...' Info
        $portName = $script:SerialPort.PortName
        Start-Sleep -Milliseconds 500
        $null = Connect-Sensor -PortName $portName -Baud $newBaud
    }
    return $ok
}

function Set-SensorUploadMode {
    param (
        [ValidateRange(0, 1)]
        [int]$Mode
    )
    $ok = Invoke-SensorCommand -Command $script:Cmd.UploadMode -Data $Mode -Label "upload mode change to $($script:UploadModes[$Mode])"
    if ($ok) { $script:AutoMode = ($Mode -eq 1) }
    return $ok
}

function Set-SensorUploadInterval {
    param (
        [ValidateRange(1, 100)]
        [int]$Value
    )
    $ok = Invoke-SensorCommand -Command $script:Cmd.UploadInterval -Data $Value -Label "upload interval change to $($Value * 100) ms"
    if ($ok) { $script:UploadInterval = $Value }
    return $ok
}

function Set-SensorLedMode {
    param (
        [ValidateRange(0, 3)]
        [int]$Mode
    )
    return Invoke-SensorCommand -Command $script:Cmd.LedMode -Data $Mode -Label "LED mode change to $($script:LedModes[$Mode])"
}

function Set-SensorRelayMode {
    param (
        [ValidateRange(0, 1)]
        [int]$Mode
    )
    return Invoke-SensorCommand -Command $script:Cmd.RelayMode -Data $Mode -Label "relay mode change to $($script:RelayModes[$Mode])"
}

function Set-SensorLineMode {
    param (
        [ValidateRange(0, 1)]
        [int]$Mode
    )
    # Datasheet: sent as 0x61 0x30, answered as 0x62 0x30
    return Invoke-SensorCommand -Header $script:LineModeHeader -Command $script:Cmd.BaudRate -Data $Mode `
        -Label "line mode change to $($script:LineModes[$Mode])"
}

function Set-SensorResponseSpeed {
    param (
        [ValidateRange(0, 2)]
        [int]$Mode
    )
    # Datasheet: sent as 0x61 0x32, answered as 0x62 0x32
    return Invoke-SensorCommand -Header $script:LineModeHeader -Command $script:Cmd.Address -Data $Mode `
        -Label "response speed change to $($script:ResponseModes[$Mode])"
}

function Watch-SensorDistance {
    param (
        [int]$Seconds = 30,
        [int]$IntervalMs = 500,
        [string]$Path
    )

    if (-not (Test-SensorConnection)) { return }

    $modeText = if ($script:AutoMode) { 'automatic upload' } else { "query every $IntervalMs ms" }
    Write-Status "Monitoring for $Seconds s ($modeText). Press Q to stop." Info
    if ($Path) { Write-Status "Writing readings to $Path" Info }

    $count = 0
    $sum = 0
    $min = [int]::MaxValue
    $max = 0
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    if ($script:AutoMode) { $script:SerialPort.DiscardInBuffer() }

    while ($stopwatch.Elapsed.TotalSeconds -lt $Seconds) {
        if (Test-StopKey) { Write-Status 'Stopped by user' Warn; break }

        $distance = Get-SensorDistance -Quiet -KeepInput
        if ($null -ne $distance) {
            $count++
            $sum += $distance
            if ($distance -lt $min) { $min = $distance }
            if ($distance -gt $max) { $max = $distance }

            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
            Write-Status ("{0}  {1,5} mm" -f $timestamp, $distance) Ok
            if ($Path) {
                [pscustomobject]@{ Timestamp = $timestamp; DistanceMm = $distance } |
                    Export-Csv -Path $Path -Append -NoTypeInformation
            }
        }

        if (-not $script:AutoMode) { Start-Sleep -Milliseconds $IntervalMs }
    }

    if ($count -gt 0) {
        Write-Status ("Monitoring complete: {0} readings, min {1} mm, max {2} mm, avg {3:N0} mm" -f $count, $min, $max, ($sum / $count)) Info
    }
    else {
        Write-Status 'Monitoring complete: no readings received' Warn
    }
}

function Test-StopKey {
    # [Console]::KeyAvailable throws in hosts without a console (e.g. ISE)
    try {
        while ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::Q -or $key.Key -eq [ConsoleKey]::Escape) { return $true }
        }
    }
    catch { }
    return $false
}

# ---------------------------------------------------------------------------
# Menu helpers
# ---------------------------------------------------------------------------

function Read-Number {
    param (
        [string]$Prompt,
        [int]$Min,
        [int]$Max,
        [Nullable[int]]$Default = $null
    )
    $text = Read-Host $Prompt
    if ([string]::IsNullOrWhiteSpace($text)) {
        if ($null -ne $Default) { return $Default }
        Write-Status 'No input' Warn
        return $null
    }

    $text = $text.Trim()
    $value = 0
    $parsed = if ($text -match '^0x[0-9a-fA-F]+$') {
        [int]::TryParse($text.Substring(2), [System.Globalization.NumberStyles]::HexNumber, $null, [ref]$value)
    }
    else {
        [int]::TryParse($text, [ref]$value)
    }

    if ($parsed -and $value -ge $Min -and $value -le $Max) { return $value }
    Write-Status "Invalid input. Allowed: $Min-$Max" Error
    return $null
}

function Read-Choice {
    param (
        [string]$Title,
        [string[]]$Options,
        [Nullable[int]]$Default = $null
    )
    Write-Status "${Title}:" Info
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Status "$i. $($Options[$i])"
    }
    $prompt = "Select (0-$($Options.Count - 1))"
    if ($null -ne $Default) { $prompt += " [default $Default]" }
    return Read-Number -Prompt $prompt -Min 0 -Max ($Options.Count - 1) -Default $Default
}

function Read-Confirmation {
    param ([string]$Question)
    $answer = Read-Host "$Question (y/n)"
    return ($answer -match '^(y|yes|j|ja)$')
}

function Show-Menu {
    Clear-Host
    Write-Status '===== XKC-KL200 Laser Sensor Management Tool v2.1 =====' Info
    Write-Host
    Write-Host 'Connection:  ' -NoNewline
    if (Test-SensorConnection -Quiet) {
        Write-Status "$($script:SerialPort.PortName) at $($script:SerialPort.BaudRate) baud" Ok
    }
    else {
        Write-Status 'Not connected' Error
    }
    $modeText = if ($script:AutoMode) { "Automatic ($($script:UploadInterval * 100) ms)" } else { 'Manual' }
    Write-Status "Upload mode: $modeText" Warn
    $distanceText = if ($null -ne $script:LastDistance) { "$($script:LastDistance) mm" } else { '-' }
    Write-Status "Last value:  $distanceText" Warn
    Write-Host
    Write-Status ' 1. Connect to sensor'
    Write-Status ' 2. Disconnect'
    Write-Status ' 3. Read distance'
    Write-Status ' 4. Monitor distance'
    Write-Status ' 5. Hard reset (factory defaults)'
    Write-Status ' 6. Soft reset (user defaults)'
    Write-Status ' 7. Change sensor address'
    Write-Status ' 8. Change baud rate'
    Write-Status ' 9. Set upload mode'
    Write-Status '10. Set upload interval'
    Write-Status '11. Set LED mode'
    Write-Status '12. Set relay mode'
    Write-Status '13. Set line output mode (relay / UART)'
    Write-Status '14. Read sensor configuration'
    Write-Status '15. Set response speed'
    Write-Status ' Q. Quit'
    Write-Host
}

function Invoke-ConnectMenu {
    $ports = @([System.IO.Ports.SerialPort]::GetPortNames() | Sort-Object)
    if ($ports.Count -eq 0) {
        Write-Status 'No serial ports found' Error
        return
    }

    $portIndex = Read-Choice -Title 'Available ports' -Options $ports -Default 0
    if ($null -eq $portIndex) { return }

    $baudOptions = $script:BaudRates | ForEach-Object { "$_ baud" }
    $baudIndex = Read-Choice -Title 'Baud rate' -Options $baudOptions -Default $script:DefaultBaudIndex
    if ($null -eq $baudIndex) { return }

    $null = Connect-Sensor -PortName $ports[$portIndex] -Baud $script:BaudRates[$baudIndex]
}

function Start-Menu {
    while ($true) {
        Show-Menu
        $choice = (Read-Host 'Enter your choice').Trim()

        if ($choice -eq 'q') {
            Write-Status 'Exiting...' Info
            return
        }

        # Everything except connect/disconnect needs an open port
        if ($choice -notin @('1', '2') -and $choice -match '^\d+$' -and -not (Test-SensorConnection)) {
            $null = Read-Host 'Press Enter to continue'
            continue
        }

        switch ($choice) {
            '1' { Invoke-ConnectMenu }
            '2' { Disconnect-Sensor }
            '3' { $null = Get-SensorDistance }
            '4' {
                $seconds = Read-Number -Prompt 'Duration in seconds [default 30]' -Min 1 -Max 86400 -Default 30
                if ($null -eq $seconds) { break }
                $ms = 500
                if (-not $script:AutoMode) {
                    $ms = Read-Number -Prompt 'Interval in ms [default 500]' -Min 50 -Max 60000 -Default 500
                    if ($null -eq $ms) { break }
                }
                $csv = Read-Host 'CSV file (empty = none)'
                Watch-SensorDistance -Seconds $seconds -IntervalMs $ms -Path $csv
            }
            '5' {
                if (Read-Confirmation 'Reset sensor to factory defaults?') { $null = Reset-Sensor -Mode Factory }
            }
            '6' {
                if (Read-Confirmation 'Reset sensor to user defaults?') { $null = Reset-Sensor -Mode User }
            }
            '7' {
                $value = Read-Number -Prompt 'New address (0-65534, decimal or 0x...)' -Min 0 -Max 65534
                if ($null -ne $value) { $null = Set-SensorAddress -NewAddress $value }
            }
            '8' {
                $baudOptions = $script:BaudRates | ForEach-Object { "$_ baud" }
                $index = Read-Choice -Title 'Baud rate' -Options $baudOptions
                if ($null -ne $index) { $null = Set-SensorBaudRate -Index $index }
            }
            '9' {
                $mode = Read-Choice -Title 'Upload mode' -Options $script:UploadModes
                if ($null -ne $mode) { $null = Set-SensorUploadMode -Mode $mode }
            }
            '10' {
                $value = Read-Number -Prompt 'Upload interval (1-100 = 100 ms - 10 s)' -Min 1 -Max 100
                if ($null -ne $value) { $null = Set-SensorUploadInterval -Value $value }
            }
            '11' {
                $mode = Read-Choice -Title 'LED mode' -Options $script:LedModes
                if ($null -ne $mode) { $null = Set-SensorLedMode -Mode $mode }
            }
            '12' {
                $mode = Read-Choice -Title 'Relay mode (only used in relay line mode)' -Options $script:RelayModes
                if ($null -ne $mode) { $null = Set-SensorRelayMode -Mode $mode }
            }
            '13' {
                $mode = Read-Choice -Title 'Line output mode' -Options $script:LineModes
                if ($null -eq $mode) { break }
                if ($mode -eq 0 -and -not (Read-Confirmation 'In relay mode the sensor sends no UART data any more. Continue?')) { break }
                $null = Set-SensorLineMode -Mode $mode
            }
            '14' { Show-SensorConfig }
            '15' {
                $mode = Read-Choice -Title 'Response speed' -Options $script:ResponseModes
                if ($null -ne $mode) { $null = Set-SensorResponseSpeed -Mode $mode }
            }
            default { Write-Status 'Invalid choice. Please try again.' Error }
        }

        $null = Read-Host 'Press Enter to continue'
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

try {
    if (-not ('System.IO.Ports.SerialPort' -as [type])) {
        Add-Type -AssemblyName System.IO.Ports
    }
}
catch {
    Write-Status "System.IO.Ports is not available: $($_.Exception.Message)" Error
    exit 1
}

try {
    if ($Port) {
        if (-not (Connect-Sensor -PortName $Port -Baud $BaudRate)) {
            if ($Action -ne 'Menu') { exit 1 }
        }
    }
    elseif ($Action -ne 'Menu') {
        Write-Status "-Action $Action requires -Port" Error
        exit 1
    }

    switch ($Action) {
        'Menu'         { Start-Menu }
        'ReadDistance' { if ($null -eq (Get-SensorDistance)) { exit 2 } }
        'ReadConfig'   { Show-SensorConfig }
        'Monitor'      { Watch-SensorDistance -Seconds $Duration -IntervalMs $Interval -Path $CsvPath }
    }
}
finally {
    # Always release the COM port, also on Ctrl+C or errors
    Disconnect-Sensor -Quiet
}
