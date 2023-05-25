$shellPath = Get-Location
[System.Environment]::SetEnvironmentVariable("CONVERT_AUDIO_ROOT", $shellPath, "Machine")
