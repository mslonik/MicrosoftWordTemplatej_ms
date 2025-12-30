
# Load Windows Forms for MsgBox
Add-Type -AssemblyName System.Windows.Forms

function Show-Message {
    param (
        [string]$Text,
        [string]$Title = "Script Notification",
        [string]$Type = "Info"
    )

    $icon = switch ($Type) {
        "Info" { [System.Windows.Forms.MessageBoxIcon]::Information }
        "Warning" { [System.Windows.Forms.MessageBoxIcon]::Warning }
        "Error" { [System.Windows.Forms.MessageBoxIcon]::Error }
        default { [System.Windows.Forms.MessageBoxIcon]::None }
    }

    [System.Windows.Forms.MessageBox]::Show($Text, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
}

# Define system paths using environment variables
$RepoBase = Join-Path $UserProfile "RepozytoriaSVN\CompanyTemplates\UnofficialTemplates"
$UserProfile = $env:USERPROFILE
$AppDataRoaming = Join-Path $UserProfile "AppData\Roaming"
$WordStartup = Join-Path $AppDataRoaming "Microsoft\Word\startup"
$Templates = Join-Path $AppDataRoaming "Microsoft\Templates"
$ThemeBase = Join-Path $Templates "Document Themes"

# Define file name variables
$RibbonCustomization = "ms_Customizations.exportedUI"
$Macros = "vSP_ms_Macros.dotm"
$BuildingBlocks = "vSP_ms_BB.dotm"
$Theme = "vSP_ms_Theme.thmx"

# Define source and destination mappings
$files = @(
    @{ Source = Join-Path $RepoBase "Normal_ConfigurationFile\Files\Normal.dotm"; Destination = Join-Path $Templates "Normal.dotm" },
    @{ Source = Join-Path $RepoBase "MenuRibbon_ConfigurationFile\Files\$RibbonCustomization"; Destination = Join-Path $WordStartup $RibbonCustomization },
    @{ Source = Join-Path $RepoBase "Macros_ConfigurationFile\Files\$Macros"; Destination = Join-Path $WordStartup $Macros },
    @{ Source = Join-Path $RepoBase "BuildingBlocks_ConfigurationFile\Files\$BuildingBlocks"; Destination = Join-Path $WordStartup $BuildingBlocks },
    @{ Source = Join-Path $RepoBase "Theme_ConfigurationFile\Files\$Theme"; Destination = Join-Path $ThemeBase $Theme }
)

# Perform conditional copy
foreach ($file in $files) {
    $src = $file.Source
    $dst = $file.Destination

    if (Test-Path $src) {
        if (!(Test-Path $dst) -or ((Get-Item $src).LastWriteTime -gt (Get-Item $dst).LastWriteTime)) {
            Copy-Item $src -Destination $dst -Force
            Show-Message "Copied:`n$src ->`n$dst" "Script Notification" "Info"
        } else {
            Show-Message "Skipped (up-to-date):`n$dst" "Script Notification" "Info"
        }
    } else {
        Show-Message "Source file not found:`n$src" "Script Notification" "Warning"
    }
}
