
function Say {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Text,
        [Parameter()]
        $Color = 'Cyan'
    )

    try {
        if ($Host.UI -and $Host.UI.RawUI) {
            $Host.UI.RawUI.ForegroundColor = $Color
        }
        $Text | Out-Host
    }
    catch {
        # Fallback: never throw from logging
        $Text
    }
    finally {
        try { [Console]::ResetColor() } catch {}
    }
}

function SayError {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Text,
        [Parameter()]
        $Color = 'Red'
    )
    try {
        if ($Host.UI -and $Host.UI.RawUI) {
            $Host.UI.RawUI.ForegroundColor = $Color
        }
        "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [ERR] - $Text" | Out-Host
    }
    catch {
        # Fallback: never throw from logging
        "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [ERR] - $Text"
    }
    finally {
        try { [Console]::ResetColor() } catch {}
    }
}


function SayInfo {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Text,
        [Parameter()]
        $Color = 'Green'
    )
    try {
        if ($Host.UI -and $Host.UI.RawUI) {
            $Host.UI.RawUI.ForegroundColor = $Color
        }
        "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [INF] - $Text" | Out-Host
    }
    catch {
        # Fallback: never throw from logging
        "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [INF] - $Text"
    }
    finally {
        try { [Console]::ResetColor() } catch {}
    }
}

function SayWarning {
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Text,
        [Parameter()]
        $Color = 'DarkYellow'
    )
    try {
        if ($Host.UI -and $Host.UI.RawUI) {
            $Host.UI.RawUI.ForegroundColor = $Color
        }
        "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [WRN] - $Text" | Out-Host
    }
    catch {
        # Fallback: never throw from logging
        "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : [WRN] - $Text"
    }
    finally {
        try { [Console]::ResetColor() } catch {}
    }
}