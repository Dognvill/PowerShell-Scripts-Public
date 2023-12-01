Add-Type -AssemblyName System.Windows.Forms

function Get-RandomPoint {
    $screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $x = Get-Random -Minimum $screen.X -Maximum ($screen.Width - $screen.X)
    $y = Get-Random -Minimum $screen.Y -Maximum ($screen.Height - $screen.Y)
    New-Object System.Drawing.Point($x, $y)
}

while ($true) {
    $randomPoint = Get-RandomPoint
    [System.Windows.Forms.Cursor]::Position = $randomPoint
    Start-Sleep -Seconds 2
}
