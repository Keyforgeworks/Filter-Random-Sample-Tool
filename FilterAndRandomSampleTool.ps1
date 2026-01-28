####################################
# Filter and Random Sample Tool
####################################

# =====================
# Splash screen
# =====================

Add-Type -AssemblyName System.Windows.Forms
$splashForm = New-Object System.Windows.Forms.Form
$splashForm.StartPosition = "CenterScreen"
$splashForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$splashForm.BackColor = [System.Drawing.SystemColors]::Control
$splashForm.Size = New-Object System.Drawing.Size(300, 150)
$splashForm.TopMost = $true
$splashForm.Text = "Filter and Random Sample Tool"
$splashForm.ShowIcon = $false

$splashLabel = New-Object System.Windows.Forms.Label
$splashLabel.Text = "Initializing...Please wait"
$splashLabel.AutoSize = $false
$splashLabel.Size = New-Object System.Drawing.Size(280, 90)
$splashLabel.Location = New-Object System.Drawing.Point(0, 5) 
$splashLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$splashLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$splashForm.Controls.Add($splashLabel)

$splashForm.Show()
$splashForm.Update()

# =====================
# Initialize
# =====================

# Set Host Log Prefs
$VerbosePreference     = 'SilentlyContinue'
$WarningPreference     = 'Continue' 
$ErrorActionPreference = 'Stop'

# Set execution policy  - we will scope to this process only - still misbehaves very rarely in the shell, not sure why yet
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Import required modules
try {
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    $splashLabel.Text = "Filter and Random Sample Tool`n`nInstalling ImportExcel module...`nPlease wait."
    $splashForm.Update()
    
    if (-not (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue)) {
        Register-PSRepository -Name 'PSGallery' `
                          -SourceLocation 'https://www.powershellgallery.com/api/v2' `
                          -InstallationPolicy 'Trusted'
    }

    Install-Module ImportExcel -Scope CurrentUser -Force
    Import-Module ImportExcel -ErrorAction Stop
}

# =====================
# Public Variables
# =====================

# App Version
$scriptVersion = "Beta_8.2"

# UI Element Sizing and Positioning
$logoScaleFactor = 0.60
$logoX = 10
$logoY = 20

# Form Dimensions
$formWidth = 900
$formHeight = 850
$FormInitHeight = 0.91
$FormInitwidth = 0.91
$FormMinH = 800
$formMinW = 600

# Sheet Dropdown Properties
$sheetDropdownWidth = 300
$sheetDropdownHeight = 30
$sheetDropdownYOffset = 30
$sheetDropdownXGap = 30

# Standard Position Values
$labelX = 20
$textX = 300
$rowHeight = 32
$tbYOffset = 4

# Column List Box Properties
$clbColumnsWidth = 250
$clbColumnsHeight = 150

# Filter Panel Properties
$panelFiltersWidth = 850
$panelFiltersHeight = 175
$panelFiltersTopPadding = 14

# Button Dimensions
$btnLoadColumnsWidth = 110
$btnGenFiltersWidth = 130
$btnSubmitWidth = 80
$btnSubmitHeight = 30
$btnClearWidth = 80
$btnClearHeight = 30

# Dynamic UI Element Properties
$textBoxWidth = 200
$dynamicLabelX = 10
$dynamicTextBoxX = 450
$dynamicRowHeight = 40

# Processing Properties
$Sleeptime = 200
$ProcLabelVerticalPosition = 0.52
$maxRows = 3     # max rows to parse for date algorithm. Pretty massive performance impact

# Color Definitions
$filteredColor = [System.Drawing.Color]::FromArgb(252, 74, 74)    		# Red
$sampleColor = [System.Drawing.Color]::FromArgb(77, 147, 217)   	# Blue
$oversampleColor = [System.Drawing.Color]::FromArgb(71, 211, 89)   # Green

# ==============================
# Custom Logging and Transcript
# ==============================

$logFolder = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Filter_Random_Sample_Tool"
if (-not (Test-Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
}

$logPath        = Join-Path -Path $logFolder -ChildPath "Log.txt"
$transcriptPath = Join-Path -Path $logFolder -ChildPath "HostTranscript.txt"

# clear logs
if (Test-Path $logPath) {
    Clear-Content -Path $logPath -ErrorAction SilentlyContinue
}

# clear tanscript
if (Test-Path $transcriptPath) {
    Remove-Item -Path $transcriptPath -Force -ErrorAction SilentlyContinue
}

Start-Transcript -Path $transcriptPath | Out-Null

function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp [$Level] $Message"
    Add-Content -Path $logPath -Value $entry
}

function Show-ErrorMessage {
    param([string]$msg)
    Write-LogMessage $msg "ERROR"
	[System.Windows.Forms.MessageBox]::Show($msg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)    
}

# ==============================
# Load assemblies and modules
# ==============================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
Import-Module ImportExcel -ErrorAction Stop

# ==============================
# Define Fonts
# ==============================

$labelFont       = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$textBoxFont     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$clbCustomFont   = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$sheetDropdownFont = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)

# ==============================
# Base 64 logo handler
# ==============================

#region base64 - Logo removed for public version
$base64Image = $null
#endregion

# Convert and scale base64 to Image
if ($base64Image) {
    try {
        $bytes = [Convert]::FromBase64String($base64Image.Trim())
        $ms = New-Object System.IO.MemoryStream(, $bytes)
        $ms.Position = 0
        $originalLogo = [System.Drawing.Image]::FromStream($ms)
        
        $newWidth  = [int]($originalLogo.Width  * $logoScaleFactor)
        $newHeight = [int]($originalLogo.Height * $logoScaleFactor)
        $logoImage = New-Object System.Drawing.Bitmap($originalLogo, $newWidth, $newHeight)
        
        if ($ms) { [void]$ms.Dispose() }
    }
    catch {
        Write-LogMessage "Error creating logo image: $($_.Exception.Message)" "ERROR"
        $logoImage = $null
    }
}

# ==============================
# Excel File Selection
# ==============================

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Title  = "Select Excel File"
$openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$openFileDialog.Multiselect = $false
if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Show-ErrorMessage "No file selected. Exiting..."
	Stop-Transcript | Out-Null
    return
}
$excelPath = $openFileDialog.FileName

# ==============================
# Function lib
# ==============================

# ---------------------
# Scaling
# ---------------------

# Store original form size for scaling calculations
$script:originalFormWidth = $formWidth
$script:originalFormHeight = $formHeight
$script:controlProperties = @{}

function Ensure-ControlHasName {
    param (
        [System.Windows.Forms.Control]$control,
        [string]$prefix
    )
    
    if ([string]::IsNullOrEmpty($control.Name)) {
        $control.Name = "$prefix" + [Guid]::NewGuid().ToString("N")
    }
    return $control
}

function Store-ControlProperties {
    param (
        [System.Windows.Forms.Control]$control
    )
    
    if (-not $script:controlProperties.ContainsKey($control.Name)) {
        $script:controlProperties[$control.Name] = @{
            X = $control.Location.X
            Y = $control.Location.Y
            Width = $control.Width
            Height = $control.Height
            FontSize = $control.Font.Size
        }
    }
    
    # Recursively store properties for child controls
    if ($control -is [System.Windows.Forms.Panel] -or $control -is [System.Windows.Forms.GroupBox]) {
        foreach ($childControl in $control.Controls) {
            if ([string]::IsNullOrEmpty($childControl.Name)) {
                $childPrefix = if ($control -is [System.Windows.Forms.Panel]) { "PanelChild_" } else { "GroupBoxChild_" }
                $childControl = Ensure-ControlHasName -control $childControl -prefix $childPrefix
            }
            Store-ControlProperties -control $childControl
        }
    }
}

function Scale-Control {
    param (
        [System.Windows.Forms.Control]$control,
        [double]$widthRatio,
        [double]$heightRatio
    )

    # Only scale controls we previously stored
    if (-not $script:controlProperties.ContainsKey($control.Name)) {
        return
    }

    # Pull original position/size
    $orig = $script:controlProperties[$control.Name]
    $newX      = [int]($orig.X      * $widthRatio)
    $newY      = [int]($orig.Y      * $heightRatio)
    $newWidth  = [int]($orig.Width  * $widthRatio)
    $newHeight = [int]($orig.Height * $heightRatio)

    # Move and resize the control
    $control.Location = New-Object System.Drawing.Point($newX, $newY)
    $control.Size     = New-Object System.Drawing.Size($newWidth, $newHeight)

    # Scale the font for everything except PictureBox
    if (-not ($control -is [System.Windows.Forms.PictureBox])) {
        # Use average of width/height ratio for balanced font scaling
        $avgRatio   = ($widthRatio + $heightRatio) / 2
        $newFontSize= $orig.FontSize * [Math]::Sqrt($avgRatio)
        $control.Font = New-Object System.Drawing.Font(
            $control.Font.FontFamily,
            $newFontSize,
            $control.Font.Style
        )
    }

    # Recurse into child containers
    if ($control -is [System.Windows.Forms.Panel] -or $control -is [System.Windows.Forms.GroupBox]) {
        foreach ($child in $control.Controls) {
            Scale-Control -control $child -widthRatio $widthRatio -heightRatio $heightRatio
        }
    }
}

# ---------------------
# Date handling
# ---------------------

function Detect-DateColumns {
    param(
        [array]$data
    )
    $dateColumns = @()
    if ($data.Count -eq 0) {
        return $dateColumns
    }

    if ($data.Count -lt $maxRows) {
        $rowsToProcess = $data
    } else {
        $rowsToProcess = $data[0..($maxRows - 1)]
    }

    # Check each property in the first row
    $firstRow = $data[0]
    foreach ($property in $firstRow.PSObject.Properties) {
        $colName        = $property.Name
        $convertedCount = 0
        $totalCount     = 0

        foreach ($row in $rowsToProcess) {
            $val = $row.$colName

            # Skip empty/whitespace
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $totalCount++
                $parsedAsDate = $false

                # Attempt to treat numeric values as OADate only in a sensible range
                if ($val -is [double]) {
                    # Only consider it as a date if it's within a range - [1927 - 2152]
                    if ($val -ge 10000 -and $val -le 70000) {
                        try {
                            [void][datetime]::FromOADate($val)
                            $parsedAsDate = $true
                        } catch { }
                    }
                }
                else {
                    # Attempt to parse strings as date formats
                    $cleanVal = $val -replace '(\d+)(st|nd|rd|th)', '$1'
                    try {
                        [void][datetime]$cleanVal
                        $parsedAsDate = $true
                    }
                    catch { }
                }

                if ($parsedAsDate) {
                    $convertedCount++
                }
            }
        }

        # If more than half the non blank rows parse as date, assume it's a date
        if ($totalCount -gt 0 -and ($convertedCount / $totalCount) -ge 0.5) {
            $dateColumns += $colName
        }
    }

    return $dateColumns
}

function Convert-DateCells {
    param(
        [array]$data,
        [string[]]$dateColumns
    )
    foreach ($row in $data) {
        foreach ($col in $dateColumns) {
            $val = $row.$col

            # Already a DateTime? Skip
            if ($val -is [DateTime]) {
                continue
            }

            # Attempt OADate within a reasonable date range
            if ($val -is [double]) {
                if ($val -ge 10000 -and $val -le 70000) {
                    try {
                        $row.$col = [datetime]::FromOADate($val)
                        continue
                    } catch { }
                }
            }

            # Otherwise, try string parse
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $cleanVal = $val -replace '(\d+)(st|nd|rd|th)', '$1'
                try {
                    $row.$col = [datetime]$cleanVal
                }
                catch { }
            }
        }
    }
}

function Convert-DateColumnsForDisplay {
    param(
        [Parameter(Mandatory)]
        [array]$dataSet,

        [Parameter(Mandatory)]
        [string[]]$dateColumns,

        [string]$format = 'MM/dd/yyyy'
    )
    
    # Create a clone to avoid modifying the original data
    $clonedData = $dataSet | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    
    foreach ($row in $clonedData) {
        foreach ($col in $dateColumns) {
            if ($row.$col -is [DateTime]) {
                $row.$col = $row.$col.ToString($format)
            }
        }
    }
    
    return $clonedData
}

# ---------------------
# Helper funcs
# ---------------------

function Add-Label($text, $x, $y) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = $text
    $lbl.Font     = $labelFont
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    $lbl = Ensure-ControlHasName -control $lbl -prefix "Label_"
    [void]$form.Controls.Add($lbl)
    Store-ControlProperties -control $lbl
    return $lbl
}

function Add-TextBox($default, $x, $y) {
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Text     = $default
    $tb.Font     = $textBoxFont
    $tb.Width    = $textBoxWidth
    $tb.Location = New-Object System.Drawing.Point($x, $y)
    $tb = Ensure-ControlHasName -control $tb -prefix "TextBox_"
    [void]$form.Controls.Add($tb)
    Store-ControlProperties -control $tb
    return $tb
}

# Gridview needs to run in a seperate thread or it breaks DPI scaling...Stupid Microsoft
function Show-GridView {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        $InputObject,
        [string] $Title = 'GridView'
    )
    begin { $buf = @() }
    process { $buf += $InputObject }
    end {
        if (-not $buf) { return }
        $tmp = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(),'xml')
        $buf | Export-CliXml -Path $tmp
        $cmd = @"
`$d = Import-CliXml '$tmp'
`$d | Out-GridView -Title '$Title' -wait
Remove-Item '$tmp'
"@ -replace "`r?`n",' ; '
        Start-Process powershell -WindowStyle Hidden -ArgumentList '-NoLogo','-NoProfile','-Command', $cmd
    }
}

function Create-ProcessingLabel {
    param (
        [string]$text = "Processing...Please wait"
    )
    
    # Remove any existing processing label first
    $existingLabel = $form.Controls | Where-Object { $_.Name -eq "processingLabel" }
    if ($existingLabel) {
        $form.Controls.Remove($existingLabel)
    }
    
    $processingLabel = New-Object System.Windows.Forms.Label
    $processingLabel.Name = "processingLabel"
    $processingLabel.Text = $text
    $processingLabel.AutoSize = $true
    $processingLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $processingLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $processingLabel.BackColor = [System.Drawing.Color]::LightYellow
    $processingLabel.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
    $processingLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    
    [void]$form.Controls.Add($processingLabel)
    
    $processingLabel.Location = New-Object System.Drawing.Point(
        [int](($form.ClientSize.Width - $processingLabel.Width) / 2),
        [int](($form.ClientSize.Height * $ProcLabelVerticalPosition) - ($processingLabel.Height / 2))
    )
    
    $processingLabel.BringToFront()
    $form.Refresh()
    
    return $processingLabel
}

function Remove-ProcessingLabel {
    $processingLabel = $form.Controls | Where-Object { $_.Name -eq "processingLabel" }
    if ($processingLabel) {
        $form.Controls.Remove($processingLabel)
        $form.Refresh()
    }
}

# ---------------------
# Excel
# ---------------------

function Export-ResultsToExcel {
    param(
        [Parameter(Mandatory)]
        [array]$filteredData,
        
        [Parameter(Mandatory)]
        [array]$randomSample,
        
        [Parameter(Mandatory)]
        [array]$randomOverSample,
        
        [Parameter(Mandatory)]
        [string]$excelPath,
        
        [Parameter(Mandatory)]
        [string[]]$dateColumns,
        
        [Parameter(Mandatory)]
        [string]$sampleFormat
    )

    try {
        # Remove any existing file
        if (Test-Path $excelPath) { Remove-Item $excelPath -Force }
        Write-LogMessage "Starting Excel export to $excelPath" "INFO"

        # Strip out any Original_* properties
        $cleanFiltered    = $filteredData    | Select-Object * -ExcludeProperty ($filteredData[0].PSObject.Properties.Name | Where-Object { $_ -like 'Original_*' })
        $cleanSample      = $randomSample     | Select-Object * -ExcludeProperty ($randomSample[0].PSObject.Properties.Name    | Where-Object { $_ -like 'Original_*' })
        $cleanOverSample  = $randomOverSample | Select-Object * -ExcludeProperty ($randomOverSample[0].PSObject.Properties.Name| Where-Object { $_ -like 'Original_*' })

        # Build dynamic labels
        $columnHeader  = 'Sample Type'

        # Force-include these columns as dates
        $forcedDates = @(
            'Escalation Option Letter',
            'QIO Verbal Notice Date',
            'Bene Verbal Notice'
        )
        $allDateCols = ($dateColumns + $forcedDates) | Select-Object -Unique

        # Convert numericstring OADates back to [datetime]
        foreach ($ds in @($cleanFiltered, $cleanSample, $cleanOverSample)) {
            foreach ($row in $ds) {
                foreach ($col in $allDateCols) {
                    $val = $row.$col
                    if ($val -is [string] -and $val -match '^\d+(\.\d+)?$') {
                        $num = [double]$val
                        if ($num -gt 10000 -and $num -lt 70000) {
                            try { $row.$col = [datetime]::FromOADate($num) } catch {}
                        }
                    }
                }
            }
        }

        # Create the Sample format by replacing the word "Sample" with "Sample {#}" if there's no {#} marker already
        $sampleFormatWithNumber = if ($sampleFormat -match '\{#\}') {
            $sampleFormat
        } else {
            if ($sampleFormat -match 'Sample') {
                $sampleFormat -replace 'Sample', 'Sample {#}'
            } else {
                # If "Sample" doesn't exist in the format, append it before any text
                "Sample {#} $sampleFormat"
            }
        }
        
        # Create the Oversample format by replacing "Sample" with "OS" if it exists
        $osFormatWithNumber = $sampleFormatWithNumber -replace 'Sample', 'OS'

        # Populate "Sample Type" in Random Sample
        $ctr = 1
        $cleanSample = $cleanSample | ForEach-Object {
            $num = $ctr.ToString('00')
            $text = $sampleFormatWithNumber -replace '\{#\}', $num
            $_ | Add-Member -MemberType NoteProperty -Name $columnHeader -Value $text -PassThru
            $ctr++
        }

        # Populate "Sample Type" in Random Oversample
        $ctr = 1
        $cleanOverSample = $cleanOverSample | ForEach-Object {
            $num = $ctr.ToString('00')
            $text = $osFormatWithNumber -replace '\{#\}', $num
            $_ | Add-Member -MemberType NoteProperty -Name $columnHeader -Value $text -PassThru
            $ctr++
        }

        # Export each sheet
        $cleanFiltered |
            Export-Excel -Path $excelPath -WorksheetName "Filtered Results" -AutoSize

        # Random Sample sheet
        $props = @($columnHeader) + (
            $cleanSample[0].PSObject.Properties.Name |
            Where-Object { $_ -ne $columnHeader }
        )
        $cleanSample |
            Select-Object -Property $props |
            Export-Excel -Path $excelPath -WorksheetName "Random Sample" -AutoSize

        # Random Oversample sheet
        $props = @($columnHeader) + (
            $cleanOverSample[0].PSObject.Properties.Name |
            Where-Object { $_ -ne $columnHeader }
        )
        $cleanOverSample |
            Select-Object -Property $props |
            Export-Excel -Path $excelPath -WorksheetName "Random Oversample" -AutoSize

        # Re-open and format dates + style headers
        $pkg = Open-ExcelPackage -Path $excelPath
        foreach ($ws in $pkg.Workbook.Worksheets) {
            $headerColor = switch ($ws.Name) {
                "Filtered Results"   { $filteredColor }
                "Random Sample"      { $sampleColor }
                "Random Oversample"  { $oversampleColor }
                default               { [System.Drawing.Color]::LightGray }
            }
            $headerRange = $ws.Cells[1,1,1,$ws.Dimension.Columns]
            $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $headerRange.Style.Fill.BackgroundColor.SetColor($headerColor)
            $headerRange.Style.Font.Bold = $true
            $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)

            for ($col = 1; $col -le $ws.Dimension.Columns; $col++) {
                $hdr = $ws.Cells[1,$col].Value
                if ($allDateCols -contains $hdr) {
                    $ws.Cells[2,$col,$ws.Dimension.Rows,$col].Style.Numberformat.Format = "mmm d, yyyy"
                }
            }

            $range = $ws.Cells[$ws.Dimension.Address]
            $range.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
            $range.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
        }

        $pkg.Save()
        $pkg.Dispose()

        Write-LogMessage "Export completed successfully" "INFO"
        return "Results successfully exported!"
    }
    catch {
        Write-LogMessage "Export error details: $($_.Exception)" "ERROR"
        throw "Error exporting to Excel: $($_.Exception.Message)"
    }
}

# ---------------------
# Filter set
# ---------------------

function Get-FilterSetsPath {
    $appDataPath = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Filter_Random_Sample_Tool"
    if (-not (Test-Path $appDataPath)) {
        try {
            New-Item -ItemType Directory -Path $appDataPath -Force | Out-Null
            Write-LogMessage "Created directory: $appDataPath" "INFO"
        }
        catch {
            Write-LogMessage "Error creating directory: $($_.Exception.Message)" "ERROR"
        }
    }
    
    $filePath = Join-Path -Path $appDataPath -ChildPath "FilterSets.json"
    Write-LogMessage "Filter sets path: $filePath" "INFO"
    return $filePath
}

function Load-FilterSets {
    $filePath = Get-FilterSetsPath
    Write-LogMessage "Attempting to load filter sets" "INFO"
    
    if (Test-Path $filePath) {
        try {
            $jsonContent = Get-Content -Path $filePath -Raw
            if ([string]::IsNullOrWhiteSpace($jsonContent)) {
                Write-LogMessage "Filter sets file is empty" "INFO"
                return @{}
            }
            
            $filterSets = $jsonContent | ConvertFrom-Json
            $result     = @{}
            
            foreach ($prop in $filterSets.PSObject.Properties) {
                $name  = $prop.Name
                $value = $prop.Value
                
                # Build the object we expect, including the four new fields
                $newValue = @{
                    Worksheet      = $value.Worksheet
                    Columns        = @()
                    FilterValues   = @{}
                    SampleSize     = $null
                    OverSampleSize = $null
                    Seed           = $null
                    SamplePrefix   = $null
                }
                
                # Columns
                if ($value.PSObject.Properties.Name -contains "Columns") {
                    foreach ($col in $value.Columns) {
                        $newValue.Columns += $col
                    }
                }
                
                # FilterValues
                if ($value.PSObject.Properties.Name -contains "FilterValues") {
                    foreach ($fv in $value.FilterValues.PSObject.Properties) {
                        $newValue.FilterValues[$fv.Name] = $fv.Value
                    }
                }
                
                # SampleSize, OverSampleSize, Seed, SamplePrefix
                if ($value.PSObject.Properties.Name -contains "SampleSize") {
                    $newValue.SampleSize = $value.SampleSize
                }
                if ($value.PSObject.Properties.Name -contains "OverSampleSize") {
                    $newValue.OverSampleSize = $value.OverSampleSize
                }
                if ($value.PSObject.Properties.Name -contains "Seed") {
                    $newValue.Seed = $value.Seed
                }
                if ($value.PSObject.Properties.Name -contains "SampleFormat") {
					$newValue.SampleFormat = $value.SampleFormat
				}
                
                $result[$name] = $newValue
            }
            
            Write-LogMessage "Filter sets loaded Succesfully" "INFO"
			return $result
        }
        catch {
            Write-LogMessage "Error loading filter sets: $($_.Exception.Message)" "ERROR"
            return @{}
        }
    }
    else {
        Write-LogMessage "Filter sets file does not exist yet" "INFO"
        return @{}
    }
}

function Save-FilterSets {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$filterSets
    )
    
    $filePath = Get-FilterSetsPath
    Write-LogMessage "Attempting to save filter sets" "INFO"
    
    try {
        $json = $filterSets | ConvertTo-Json -Depth 10
        $json | Set-Content -Path $filePath -Force
        
        # Verify the file was created and has content
        if (Test-Path $filePath) {
            $fileSize = (Get-Item $filePath).Length
            Write-LogMessage "Filter sets file created, size: $fileSize bytes" "INFO"
            
            if ($fileSize -eq 0) {
                Write-LogMessage "Warning: Filter sets file is empty" "WARNING"
                return $false
            }
            
            return $true
        } else {
            Write-LogMessage "Failed to create filter sets file" "ERROR"
            return $false
        }
    }
    catch {
        Write-LogMessage "Error saving filter sets: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Update-FilterSetsList {
    Write-LogMessage "Updating filter sets dropdown" "INFO"
    $cbFilterSets.Items.Clear()
    $filterSets = Load-FilterSets
    
    if ($filterSets.Keys.Count -gt 0) {
        # Sort the keys alphabetically before adding them to the dropdown
        $sortedKeys = $filterSets.Keys | Sort-Object
        
        foreach ($name in $sortedKeys) {
            [void]$cbFilterSets.Items.Add($name)
        }
        
        if ($cbFilterSets.Items.Count -gt 0) {
            $cbFilterSets.SelectedIndex = 0
        }
    }
}

function Generate-FilterInputs {
    $yOffset = $panelFiltersTopPadding
    $panelFilters.Controls.Clear()

    foreach ($field in $clbColumns.CheckedItems) {
        $isDateColumn = $global:excelDateCols -contains $field

        if ($isDateColumn) {
            # date range
            $lblFrom = New-Object System.Windows.Forms.Label
            $lblFrom.Text     = "$field Range From:"
            $lblFrom.AutoSize = $true
            $lblFrom.Location = New-Object System.Drawing.Point($dynamicLabelX, $yOffset)
            $lblFrom = Ensure-ControlHasName -control $lblFrom -prefix "LabelFrom_"
            [void]$panelFilters.Controls.Add($lblFrom)
            Store-ControlProperties -control $lblFrom

            $tbFrom = New-Object System.Windows.Forms.TextBox
            $tbFrom.Name     = "tb_${field}_From"
            $tbFrom.Width    = $textBoxWidth
            $tbFrom.Text     = "" 
            $tbFrom.Location = New-Object System.Drawing.Point($dynamicTextBoxX, $yOffset)			
            $tbFrom = Ensure-ControlHasName -control $tbFrom -prefix "TextBoxFrom_"
            [void]$panelFilters.Controls.Add($tbFrom)
            Store-ControlProperties -control $tbFrom
            $yOffset += $dynamicRowHeight

            $lblTo = New-Object System.Windows.Forms.Label
            $lblTo.Text     = "$field Range To:"
            $lblTo.AutoSize = $true
            $lblTo.Location = New-Object System.Drawing.Point($dynamicLabelX, $yOffset)
            $lblTo = Ensure-ControlHasName -control $lblTo -prefix "LabelTo_"
            [void]$panelFilters.Controls.Add($lblTo)
            Store-ControlProperties -control $lblTo

            $tbTo = New-Object System.Windows.Forms.TextBox
            $tbTo.Name     = "tb_${field}_To"
            $tbTo.Width    = $textBoxWidth
            $tbTo.Text     = "" 
            $tbTo.Location = New-Object System.Drawing.Point($dynamicTextBoxX, $yOffset)
            $tbTo = Ensure-ControlHasName -control $tbTo -prefix "TextBoxTo_"
            [void]$panelFilters.Controls.Add($tbTo)
            Store-ControlProperties -control $tbTo
            $yOffset += $dynamicRowHeight
        }
        else {
            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text = "$($field):"
            $lbl.AutoSize = $true
            $lbl.Location = New-Object System.Drawing.Point($dynamicLabelX, $yOffset)
            $lbl = Ensure-ControlHasName -control $lbl -prefix "Label_"
            [void]$panelFilters.Controls.Add($lbl)
            Store-ControlProperties -control $lbl

            $tb = New-Object System.Windows.Forms.TextBox
            $tb.Name     = "tb_$field"
            $tb.Width    = $textBoxWidth
            $tb.Text     = "" 
            $tb.Location = New-Object System.Drawing.Point($dynamicTextBoxX, $yOffset)
            $tb = Ensure-ControlHasName -control $tb -prefix "TextBox_"
            [void]$panelFilters.Controls.Add($tb)
            Store-ControlProperties -control $tb
            $yOffset += $dynamicRowHeight
        }
    }
}

# ==============================
# Build form with scaling support
# ==============================

$form = New-Object System.Windows.Forms.Form
$form.Text          = "Filter and Random Sample Tool " + $scriptVersion
$form.Width         = $formWidth
$form.Height        = $formHeight
$form.StartPosition = "CenterScreen"
$form.TopMost       = $true
$form.MinimumSize   = New-Object System.Drawing.Size($FormMinH, $formMinW)
$form.Name          = "mainForm"

# Force the form to front, then disable TopMost
$form.Add_Shown({
    $this.Activate()
    Start-Sleep -Milliseconds $Sleeptime
    $this.TopMost = $false
})

# Add Resize event handler
$form.Add_Resize({
    if ($form.WindowState -ne [System.Windows.Forms.FormWindowState]::Minimized) {
        $currentWidth = $form.Width
        $currentHeight = $form.Height
        
        # Calculate scaling factors
        $widthRatio = $currentWidth / $script:originalFormWidth
        $heightRatio = $currentHeight / $script:originalFormHeight
        
        # Scale all controls
        foreach ($control in $form.Controls) {
            Scale-Control -control $control -widthRatio $widthRatio -heightRatio $heightRatio
        }
    }
})

# Add Resize event handler
$form.Add_Resize({
    if ($form.WindowState -ne [System.Windows.Forms.FormWindowState]::Minimized) {
        $currentWidth = $form.Width
        $currentHeight = $form.Height
        
        # Calculate scaling factors
        $widthRatio = $currentWidth / $script:originalFormWidth
        $heightRatio = $currentHeight / $script:originalFormHeight
        
        # Scale all controls
        foreach ($control in $form.Controls) {
            Scale-Control -control $control -widthRatio $widthRatio -heightRatio $heightRatio
        }
    }
})

# If we have a valid $logoImage, add it to the form
if ($logoImage) {
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Image    = $logoImage
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $pictureBox.Location = New-Object System.Drawing.Point($logoX, $logoY)
    $pictureBox.Size     = New-Object System.Drawing.Size($logoImage.Width, $logoImage.Height)
    $pictureBox = Ensure-ControlHasName -control $pictureBox -prefix "Logo_"
    [void]$form.Controls.Add($pictureBox)
    Store-ControlProperties -control $pictureBox
	Write-LogMessage "Logo Image is Valid" "INFO"
}

# ===============================
# Dynamic Control Handlers
# ===============================

$currentY = $logoY
$currentY += $logoImage.Height + 10

# Worksheet ComboBox
$worksheetLabelX = $textX
$worksheetLabel  = Add-Label "Select Worksheet:" $worksheetLabelX $currentY
$comboY          = [int]$currentY + 25
$cbSheets        = New-Object System.Windows.Forms.ComboBox
$cbSheets.Location      = New-Object System.Drawing.Point($textX, $comboY)
$cbSheets.Width         = $sheetDropdownWidth
$cbSheets.Height        = $sheetDropdownHeight
$cbSheets.Font          = $sheetDropdownFont
$cbSheets.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cbSheets.Name          = "cbSheets"
[void]$form.Controls.Add($cbSheets)
Store-ControlProperties -control $cbSheets
$currentY += $rowHeight + 15

try {
    $sheetInfo  = Get-ExcelSheetInfo -Path $excelPath
    $sheetNames = $sheetInfo | ForEach-Object { $_.Name }
    $cbSheets.Items.AddRange($sheetNames)
    if ($cbSheets.Items.Count -gt 0) {
        $cbSheets.SelectedIndex = 0
    }
    Write-LogMessage "Successfully retrieved sheet names" "INFO"
}
catch {
    Show-ErrorMessage "Error retrieving sheet names: $($_.Exception.Message)"
    return
}

# Filter Set Management
$grpFilterSets       = New-Object System.Windows.Forms.GroupBox
$grpFilterSets.Text  = "Filter Set Management"
$filterSetX          = $textX
$filterSetY          = $comboY + $sheetDropdownHeight + 20
$grpFilterSets.Location = New-Object System.Drawing.Point($filterSetX, $filterSetY)
$grpFilterSets.Size     = New-Object System.Drawing.Size(250, 155)
$grpFilterSets.Name     = "grpFilterSets"
[void]$form.Controls.Add($grpFilterSets)
Store-ControlProperties -control $grpFilterSets

$btnLoadFilterSet       = New-Object System.Windows.Forms.Button
$btnLoadFilterSet.Text  = "Load Filter Set"
$btnLoadFilterSet.Location = New-Object System.Drawing.Point(15, 60)
$btnLoadFilterSet.Size     = New-Object System.Drawing.Size(100, 25)
$btnLoadFilterSet.Name     = "btnLoadFilterSet"
[void]$grpFilterSets.Controls.Add($btnLoadFilterSet)
Store-ControlProperties -control $btnLoadFilterSet

$btnSaveFilterSet       = New-Object System.Windows.Forms.Button
$btnSaveFilterSet.Text  = "Save Filter set"
$btnSaveFilterSet.Location = New-Object System.Drawing.Point(135, 60)
$btnSaveFilterSet.Size     = New-Object System.Drawing.Size(100, 25)
$btnSaveFilterSet.Name     = "btnSaveFilterSet"
[void]$grpFilterSets.Controls.Add($btnSaveFilterSet)
Store-ControlProperties -control $btnSaveFilterSet

$btnDeleteFilterSet       = New-Object System.Windows.Forms.Button
$btnDeleteFilterSet.Text  = "Delete Filter Set"
$btnDeleteFilterSet.Location = New-Object System.Drawing.Point(75, 95)
$btnDeleteFilterSet.Size     = New-Object System.Drawing.Size(100, 25)
$btnDeleteFilterSet.Name     = "btnDeleteFilterSet"
[void]$grpFilterSets.Controls.Add($btnDeleteFilterSet)
Store-ControlProperties -control $btnDeleteFilterSet

$cbFilterSets                    = New-Object System.Windows.Forms.ComboBox
$cbFilterSets.Location           = New-Object System.Drawing.Point(15, 25)
$cbFilterSets.Width              = 220
$cbFilterSets.DropDownStyle      = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cbFilterSets.Name               = "cbFilterSets"
[void]$grpFilterSets.Controls.Add($cbFilterSets)
Store-ControlProperties -control $cbFilterSets

# Populate filter-set list
Update-FilterSetsList

# Load Columns button
$btnLoadColumns       = New-Object System.Windows.Forms.Button
$btnLoadColumns.Text  = "Load Columns"
$btnLoadColumns.Width = $btnLoadColumnsWidth
$btnLoadColumns.Location = New-Object System.Drawing.Point($labelX, ($currentY - 30))
$btnLoadColumns.Name     = "btnLoadColumns"
[void]$form.Controls.Add($btnLoadColumns)
Store-ControlProperties -control $btnLoadColumns
$currentY += $rowHeight

# CheckedListBox of columns
$clbColumns                = New-Object System.Windows.Forms.CheckedListBox
$clbColumns.Location       = New-Object System.Drawing.Point($labelX, $currentY)
$clbColumns.Width          = $clbColumnsWidth
$clbColumns.Height         = $clbColumnsHeight
$clbColumns.Font           = $clbCustomFont
$clbColumns.DrawMode       = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
$clbColumns.ItemHeight     = 30
$clbColumns.IntegralHeight = $false
$clbColumns.Name           = "clbColumns"
$clbColumns.Add_DrawItem({
    param($sender, $e)
    if ($e.Index -lt 0) { return }
    $itemText = $sender.Items[$e.Index].ToString()
    if (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected) {
        $e.Graphics.FillRectangle([System.Drawing.SystemBrushes]::Highlight, $e.Bounds)
        $brush = [System.Drawing.SystemBrushes]::HighlightText
    }
    else {
        $e.Graphics.FillRectangle([System.Drawing.SystemBrushes]::Window, $e.Bounds)
        $brush = [System.Drawing.SystemBrushes]::WindowText
    }
    $e.Graphics.DrawString($itemText, $sender.Font, $brush, $e.Bounds)
    $e.DrawFocusRectangle()
})
[void]$form.Controls.Add($clbColumns)
Store-ControlProperties -control $clbColumns

$lblColumns = Add-Label "Select Fields to Filter:" $labelX ($currentY - 30)
$lblColumns.Font = $labelFont
$lblColumns.Font = $labelFont
$currentY += $clbColumnsHeight + 10

# Generate Filter Inputs button
$btnGenFilters       = New-Object System.Windows.Forms.Button
$btnGenFilters.Text  = "Generate Filter Inputs"
$btnGenFilters.Width = $btnGenFiltersWidth
$btnGenFilters.Location = New-Object System.Drawing.Point($labelX, $currentY)
$btnGenFilters.Name     = "btnGenFilters"
[void]$form.Controls.Add($btnGenFilters)
Store-ControlProperties -control $btnGenFilters
$currentY += $rowHeight

# Panel for dynamic filter controls
$panelFilters             = New-Object System.Windows.Forms.Panel
$panelFilters.Location    = New-Object System.Drawing.Point($labelX, $currentY)
$panelFilters.Size        = New-Object System.Drawing.Size($panelFiltersWidth, $panelFiltersHeight)
$panelFilters.AutoScroll  = $true
$panelFilters.AutoScrollMargin = New-Object System.Drawing.Size(0,10)
$panelFilters.Padding         = New-Object System.Windows.Forms.Padding(0,0,0,10)
$panelFilters.BorderStyle     = "FixedSingle"
$panelFilters.Name            = "panelFilters"
[void]$form.Controls.Add($panelFilters)
Store-ControlProperties -control $panelFilters
$currentY += $panelFiltersHeight + 10

# Sample Size, OverSample, Seed, Prefix
$lblSampleSize      = Add-Label "Random Sample Size:" $labelX $currentY
$txtSampleSize      = Add-TextBox "" $textX $currentY;   $txtSampleSize.Name = "txtSampleSize";   Store-ControlProperties -control $txtSampleSize
$currentY += $rowHeight

$lblOverSampleSize  = Add-Label "OverSample Size:"    $labelX $currentY
$txtOverSampleSize  = Add-TextBox "" $textX $currentY;   $txtOverSampleSize.Name = "txtOverSampleSize"; Store-ControlProperties -control $txtOverSampleSize
$currentY += $rowHeight

$lblSeed            = Add-Label "Seed:"               $labelX $currentY
$txtSeed            = Add-TextBox "" $textX $currentY;   $txtSeed.Name = "txtSeed";                 Store-ControlProperties -control $txtSeed
$currentY += $rowHeight

$lblSamplePrefix = Add-Label "Sample Format (use {#} for numbering):" $labelX $currentY
$txtSamplePrefix = Add-TextBox "" $textX $currentY; $txtSamplePrefix.Name = "txtSamplePrefix"; Store-ControlProperties -control $txtSamplePrefix
$currentY += $rowHeight

# Output Options Panel
$grpOutputOptions       = New-Object System.Windows.Forms.GroupBox
$grpOutputOptions.Text  = "Output Options"
$grpOutputOptions.Location = New-Object System.Drawing.Point($labelX, $currentY)
$grpOutputOptions.Size     = New-Object System.Drawing.Size(520, 75)
$grpOutputOptions.Name     = "grpOutputOptions"
[void]$form.Controls.Add($grpOutputOptions)
Store-ControlProperties -control $grpOutputOptions

$rbExcel               = New-Object System.Windows.Forms.RadioButton
$rbExcel.Text          = "Export to Excel File"
$rbExcel.Location      = New-Object System.Drawing.Point(15, 20)
$rbExcel.Size          = New-Object System.Drawing.Size(150, 20)
$rbExcel.Checked       = $true
$rbExcel.Name          = "rbExcel"
[void]$grpOutputOptions.Controls.Add($rbExcel)
Store-ControlProperties -control $rbExcel

$rbGridView               = New-Object System.Windows.Forms.RadioButton
$rbGridView.Text          = "Show in GridView"
$rbGridView.Location      = New-Object System.Drawing.Point(165, 20)
$rbGridView.Size          = New-Object System.Drawing.Size(150, 20)
$rbGridView.Checked       = $false
$rbGridView.Name          = "rbGridView"
[void]$grpOutputOptions.Controls.Add($rbGridView)
Store-ControlProperties -control $rbGridView

$txtExcelPath               = New-Object System.Windows.Forms.TextBox
$txtExcelPath.Location      = New-Object System.Drawing.Point(15, 45)
$txtExcelPath.Size          = New-Object System.Drawing.Size(360, 20)
$txtExcelPath.Enabled       = $false
$txtExcelPath.Name          = "txtExcelPath"
[void]$grpOutputOptions.Controls.Add($txtExcelPath)
Store-ControlProperties -control $txtExcelPath

# Script directory & default filename
$scriptDirectory = if ($PSScriptRoot -and $PSScriptRoot -ne "") {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    (Get-Location).Path
}
$defaultExcelName        = "Filter and Random Sample $(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
$txtExcelPath.Text       = Join-Path -Path $scriptDirectory -ChildPath $defaultExcelName
$txtExcelPath.Enabled    = $true

# Browse button
$btnBrowse               = New-Object System.Windows.Forms.Button
$btnBrowse.Text          = "Browse..."
$btnBrowse.Location      = New-Object System.Drawing.Point(380, 44)
$btnBrowse.Size          = New-Object System.Drawing.Size(80, 22)
$btnBrowse.Enabled       = $true
$btnBrowse.Name          = "btnBrowse"
[void]$grpOutputOptions.Controls.Add($btnBrowse)
Store-ControlProperties -control $btnBrowse

# Radio-button toggles
$rbExcel.Add_CheckedChanged({
    $txtExcelPath.Enabled = $true
    $btnBrowse.Enabled    = $true
})
$rbGridView.Add_CheckedChanged({
    $txtExcelPath.Enabled = $false
    $btnBrowse.Enabled    = $false
})

# Browse click handler
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter           = "Excel Files (*.xlsx)|*.xlsx"
    $dlg.Title            = "Save Results As"
    $dlg.FileName         = [IO.Path]::GetFileName($txtExcelPath.Text)
    $dlg.InitialDirectory = [IO.Path]::GetDirectoryName($txtExcelPath.Text)
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtExcelPath.Text = $dlg.FileName
    }
})

# Move down for Submit/Clear
$currentY += 85

# Cast X positions to int to avoid array + int errors
$openFileX = [int]$labelX
$submitX = $openFileX + [int]$btnClearWidth + 20  # Using btnClearWidth for consistency
$clearX  = $submitX + [int]$btnSubmitWidth + 20

# Open File button
$btnOpenFile = New-Object System.Windows.Forms.Button
$btnOpenFile.Text = "Open"
$btnOpenFile.Size = New-Object System.Drawing.Size($btnClearWidth, $btnClearHeight)
$btnOpenFile.Location = New-Object System.Drawing.Point($openFileX, $currentY)
$btnOpenFile.Name = "btnOpenFile"
$btnOpenFile.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$btnOpenFile.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$btnOpenFile.TabStop = $false
[void]$form.Controls.Add($btnOpenFile)
Store-ControlProperties -control $btnOpenFile

# Submit Button
$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Text = "Submit"
$btnSubmit.Size = New-Object System.Drawing.Size($btnSubmitWidth, $btnSubmitHeight)
$btnSubmit.Location = New-Object System.Drawing.Point($submitX, $currentY)
$btnSubmit.Name = "btnSubmit"
$btnSubmit.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$btnSubmit.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
[void]$form.Controls.Add($btnSubmit)
Store-ControlProperties -control $btnSubmit

# Clear Button
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Size = New-Object System.Drawing.Size($btnClearWidth, $btnClearHeight)
$btnClear.Location = New-Object System.Drawing.Point($clearX, $currentY)
$btnClear.Name = "btnClear"
$btnClear.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$btnClear.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
[void]$form.Controls.Add($btnClear)
Store-ControlProperties -control $btnClear

# Advance Y for anything after
$currentY += $rowHeight

# ===============================
# Event Handlers
# ===============================

# Global variable to store data set
$global:sampleData    = $null
$global:excelDateCols = @()

# ---------------------
# Load columns click
# ---------------------

$btnLoadColumns.Add_Click({
    # Disable all controls while processing
    foreach ($ctrl in $form.Controls) { 
        if ($ctrl -ne $processingLabel) {
            $ctrl.Enabled = $false 
        }
    }
    
    $processingLabel = Create-ProcessingLabel
    
    try {
        $selectedSheet = $cbSheets.SelectedItem
        if (-not $selectedSheet) {
            Show-ErrorMessage "Please select a worksheet."
            return
        }

		# Open the workbook and grab the named sheet
		$pkg       = Open-ExcelPackage -Path $excelPath
		$ws        = $pkg.Workbook.Worksheets[$selectedSheet]

		# Figure out where the data lives
		$startRow  = $ws.Dimension.Start.Row
		$endRow    = $ws.Dimension.End.Row
		$colCount  = $ws.Dimension.End.Column

		# Decide how many non-empty cells counts as “header” (50% of columns)
		$minHeaders = [math]::Ceiling($colCount * 0.5)
		$headerRow  = $null

		for ($r = $startRow; $r -le $endRow; $r++) {
			$vals = 1..$colCount | ForEach-Object { $ws.Cells[$r, $_].Text }
			$nonEmpty = ($vals | Where-Object { $_ -and $_.Trim() }).Count
			if ($nonEmpty -ge $minHeaders) {
				$headerRow = $r
				break
			}
		}
		if (-not $headerRow) { $headerRow = $startRow }  # fallback

		Write-Verbose "Detected header row: $headerRow"

		# Grab data for later
        
		#$dataTemp = Import-Excel -Path $excelPath -WorksheetName $selectedSheet -ErrorAction Stop
		$dataTemp = Import-Excel `
			-Path          $excelPath `
			-WorksheetName $selectedSheet `
			-HeaderRow     $headerRow `
			-ErrorAction   Stop
	   
	   # Store the header row 
		$script:headerRow = $headerRow
	   
	   if ($dataTemp.Count -eq 0) {
            Show-ErrorMessage "Selected sheet is empty."
            return
        }
        $global:sampleData = $dataTemp

        # Detect date columns and convert them
        $detectedDates = Detect-DateColumns $dataTemp
        Convert-DateCells -data $global:sampleData -dateColumns $detectedDates
        $global:excelDateCols = $detectedDates

        # Populate the checked-list with column names
        $columns = $dataTemp[0].PSObject.Properties.Name
        $clbColumns.Items.Clear()
        $clbColumns.Items.AddRange($columns)
		Write-LogMessage "Succesfully loaded columns" "INFO"
    }
    catch {
        Write-LogMessage "Error loading columns: $($_.Exception.Message)" "ERROR"
		Show-ErrorMessage "Error loading columns: $($_.Exception.Message)"
    }
    finally {
        Remove-ProcessingLabel
        # Re-enable all controls
        foreach ($ctrl in $form.Controls) { $ctrl.Enabled = $true }
    }
})

# ---------------------
# Generate filter click
# ---------------------

$btnGenFilters.Add_Click({
    # Disable all controls while processing
    foreach ($ctrl in $form.Controls) { 
        if ($ctrl -ne $processingLabel) {
            $ctrl.Enabled = $false 
        }
    }
    
    $processingLabel = Create-ProcessingLabel 
    
    try {
        Generate-FilterInputs
    }
    catch {
        Show-ErrorMessage "Error generating filters: $($_.Exception.Message)"
    }
    finally {
        Remove-ProcessingLabel
        # Re-enable all controls
        foreach ($ctrl in $form.Controls) { $ctrl.Enabled = $true }
    }
})

# ---------------------
# Save filter click
# ---------------------

$btnSaveFilterSet.Add_Click({
    # First check if we have columns selected and filters generated
    if ($clbColumns.CheckedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select columns and generate filters first.",
            "Cannot Save Filter Set",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # Check if filter input fields have been generated
    if ($panelFilters.Controls.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please generate filter inputs first.",
            "Cannot Save Filter Set",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # Prompt for filter set name
    $filterSetName = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Enter a name for this filter set:",
        "Save Filter Set",
        ""
    )
    
    # Require a non-empty name
    if ([string]::IsNullOrWhiteSpace($filterSetName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Filter set name cannot be empty.",
            "Cannot Save Filter Set",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # Collect the current filter configuration
    $filterConfig = @{
        Worksheet    = $cbSheets.SelectedItem
        Columns      = @($clbColumns.CheckedItems)
        FilterValues = @{}
    }
    
    # Collect values from all filter input textboxes
    foreach ($ctrl in $panelFilters.Controls) {
        if ($ctrl -is [System.Windows.Forms.TextBox]) {
            $filterConfig.FilterValues[$ctrl.Name] = $ctrl.Text
        }
    }

    # Also save our sampling controls
    $filterConfig.SampleSize     = $txtSampleSize.Text
    $filterConfig.OverSampleSize = $txtOverSampleSize.Text
    $filterConfig.Seed           = $txtSeed.Text
   $filterConfig.SampleFormat = $txtSamplePrefix.Text
    
    # Load existing filter sets, add/update the new one, and save
    $filterSets = Load-FilterSets
    $filterSets[$filterSetName] = $filterConfig
    
    # Save the updated filter sets
    if (Save-FilterSets -filterSets $filterSets) {
        [System.Windows.Forms.MessageBox]::Show(
            "Filter set saved successfully.",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        Update-FilterSetsList
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to save filter set.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# ---------------------
# Load filter click
# ---------------------

$btnLoadFilterSet.Add_Click({
    if ($cbFilterSets.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a filter set to load.",
            "Cannot Load Filter Set",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $filterSetName = $cbFilterSets.SelectedItem.ToString()
    Write-LogMessage "Loading filter set: $filterSetName" "INFO"
    
    $filterSets   = Load-FilterSets
    $filterConfig = $filterSets[$filterSetName]
    
    if ($filterConfig -eq $null) {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not find the selected filter set.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    
    # Set the worksheet
    $worksheetIndex = $cbSheets.Items.IndexOf($filterConfig.Worksheet)
    if ($worksheetIndex -ge 0) {
        $cbSheets.SelectedIndex = $worksheetIndex
        
        # Load the columns for this worksheet
        $btnLoadColumns.PerformClick()
        
        # Check the columns that were saved in this filter set
        foreach ($column in $filterConfig.Columns) {
            $columnIndex = $clbColumns.Items.IndexOf($column)
            if ($columnIndex -ge 0) {
                $clbColumns.SetItemChecked($columnIndex, $true)
            }
        }
        
        # Clear existing controls
        $panelFilters.Controls.Clear()
        $yOffset = $panelFiltersTopPadding
        
        # Directly create controls with values
        foreach ($field in $clbColumns.CheckedItems) {
            $isDateColumn = $global:excelDateCols -contains $field
            
            if ($isDateColumn) {
                # date range
                $lblFrom = New-Object System.Windows.Forms.Label
                $lblFrom.Text     = "$field Range From:"
                $lblFrom.AutoSize = $true
                $lblFrom.Location = New-Object System.Drawing.Point($dynamicLabelX, $yOffset)
                $lblFrom = Ensure-ControlHasName -control $lblFrom -prefix "LabelFrom_"
                [void]$panelFilters.Controls.Add($lblFrom)
                Store-ControlProperties -control $lblFrom

                $fromName  = "tb_${field}_From"
                $fromValue = if ($filterConfig.FilterValues.ContainsKey($fromName)) {
                    $filterConfig.FilterValues[$fromName]
                } else { "" }
                
                $tbFrom = New-Object System.Windows.Forms.TextBox
                $tbFrom.Name     = $fromName
                $tbFrom.Width    = $textBoxWidth
                $tbFrom.Text     = $fromValue
                $tbFrom.Location = New-Object System.Drawing.Point($dynamicTextBoxX, $yOffset)
                $tbFrom = Ensure-ControlHasName -control $tbFrom -prefix "TextBoxFrom_"
                [void]$panelFilters.Controls.Add($tbFrom)
                Store-ControlProperties -control $tbFrom
                $yOffset += $dynamicRowHeight

                $lblTo = New-Object System.Windows.Forms.Label
                $lblTo.Text     = "$field Range To:"
                $lblTo.AutoSize = $true
                $lblTo.Location = New-Object System.Drawing.Point($dynamicLabelX, $yOffset)
                $lblTo = Ensure-ControlHasName -control $lblTo -prefix "LabelTo_"
                [void]$panelFilters.Controls.Add($lblTo)
                Store-ControlProperties -control $lblTo

                $toName  = "tb_${field}_To"
                $toValue = if ($filterConfig.FilterValues.ContainsKey($toName)) {
                    $filterConfig.FilterValues[$toName]
                } else { "" }
                
                $tbTo = New-Object System.Windows.Forms.TextBox
                $tbTo.Name     = $toName
                $tbTo.Width    = $textBoxWidth
                $tbTo.Text     = $toValue
                $tbTo.Location = New-Object System.Drawing.Point($dynamicTextBoxX, $yOffset)
                $tbTo = Ensure-ControlHasName -control $tbTo -prefix "TextBoxTo_"
                [void]$panelFilters.Controls.Add($tbTo)
                Store-ControlProperties -control $tbTo
                $yOffset += $dynamicRowHeight
            }
            else {
                $lbl = New-Object System.Windows.Forms.Label
                $lbl.Text     = "$($field):"
                $lbl.AutoSize = $true
                $lbl.Location = New-Object System.Drawing.Point($dynamicLabelX, $yOffset)
                $lbl = Ensure-ControlHasName -control $lbl -prefix "Label_"
                [void]$panelFilters.Controls.Add($lbl)
                Store-ControlProperties -control $lbl

                $fieldName  = "tb_$field"
                $fieldValue = if ($filterConfig.FilterValues.ContainsKey($fieldName)) {
                    $filterConfig.FilterValues[$fieldName]
                } else { "" }
                
                $tb = New-Object System.Windows.Forms.TextBox
                $tb.Name     = $fieldName
                $tb.Width    = $textBoxWidth
                $tb.Text     = $fieldValue
                $tb.Location = New-Object System.Drawing.Point($dynamicTextBoxX, $yOffset)
                $tb = Ensure-ControlHasName -control $tb -prefix "TextBox_"
                [void]$panelFilters.Controls.Add($tb)
                Store-ControlProperties -control $tb
                $yOffset += $dynamicRowHeight
            }
        }

        # restore sampling controls
        $txtSampleSize.Text     = $filterConfig.SampleSize
        $txtOverSampleSize.Text = $filterConfig.OverSampleSize
        $txtSeed.Text           = $filterConfig.Seed
        $txtSamplePrefix.Text = $filterConfig.SampleFormat
        
        $form.Refresh()
        [System.Windows.Forms.MessageBox]::Show(
            "Filter set loaded successfully.",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "The worksheet specified in this filter set could not be found.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# ---------------------
# Delete filter click
# ---------------------

$btnDeleteFilterSet.Add_Click({
    if ($cbFilterSets.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please select a filter set to delete.", "Cannot Delete Filter Set", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $filterSetName = $cbFilterSets.SelectedItem.ToString()
    $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the filter set '$filterSetName'?", "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $filterSets = Load-FilterSets
        
        # Remove the key from the hashtable
        $filterSets.Remove($filterSetName)
        
        # Save the updated filter sets
        if (Save-FilterSets -filterSets $filterSets) {
            [System.Windows.Forms.MessageBox]::Show("Filter set deleted successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Update-FilterSetsList
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Failed to delete filter set.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# ---------------------
# Open file click
# ---------------------

$btnOpenFile.Add_Click({
    # Create a file open dialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select Excel File"
    $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    $openFileDialog.Multiselect = $false
    
    # Show the dialog and process if OK was clicked
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newExcelPath = $openFileDialog.FileName
        Write-LogMessage "Opening new Excel file: $newExcelPath" "INFO"
        
        # Create processing label
        $processingLabel = Create-ProcessingLabel "Loading file...Please wait"
        
        try {
            # Update the global variable
            $script:excelPath = $newExcelPath
            
            # Clear the current workbook data
            $cbSheets.Items.Clear()
            $clbColumns.Items.Clear()
            $panelFilters.Controls.Clear()
            $global:sampleData = $null
            $global:excelDateCols = @()
            
            # Reset other fields
            $txtSampleSize.Text = ""
            $txtOverSampleSize.Text = ""
            $txtSeed.Text = ""
            
            # Load the sheet names from the new file
            $sheetInfo = Get-ExcelSheetInfo -Path $newExcelPath
            $sheetNames = $sheetInfo | ForEach-Object { $_.Name }
            $cbSheets.Items.AddRange($sheetNames)
            if ($cbSheets.Items.Count -gt 0) {
                $cbSheets.SelectedIndex = 0
            }
            
            # Update form title to show the new file name
            $form.Text = "Filter and Random Sample Tool " + $scriptVersion + " - " + [System.IO.Path]::GetFileName($newExcelPath)
            
            # Update the Excel path in the export field
            $defaultExcelName = "Filter and Random Sample $(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            $txtExcelPath.Text = Join-Path -Path (Split-Path -Parent $newExcelPath) -ChildPath $defaultExcelName
            
            Write-LogMessage "Successfully loaded new Excel file" "INFO"
            [System.Windows.Forms.MessageBox]::Show(
                "New Excel file loaded successfully.",
                "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            Write-LogMessage "Error loading new Excel file: $($_.Exception.Message)" "ERROR"
            Show-ErrorMessage "Error loading new Excel file: $($_.Exception.Message)"
        }
        finally {
            # Remove the processing label
            Remove-ProcessingLabel
        }
    }
})

# ---------------------
# Submit click
# ---------------------

$btnSubmit.Add_Click({
    # Disable all controls while processing
    foreach ($ctrl in $form.Controls) { 
        if ($ctrl -ne $processingLabel) {
            $ctrl.Enabled = $false 
        }
    }
    
    $processingLabel = Create-ProcessingLabel 
    
    try {
        Write-LogMessage "Starting filter/sample processing" "INFO"
		if ($txtSeed.Text -and [int]::TryParse($txtSeed.Text, [ref] $null)) {
            Get-Random -SetSeed ([int]$txtSeed.Text) | Out-Null
        }
        $selectedSheet = $cbSheets.SelectedItem
        if (-not $selectedSheet) { throw "No worksheet selected." }

        # Re-import fresh data
        
		#$data = Import-Excel -Path $excelPath -WorksheetName $selectedSheet
        $data = Import-Excel `
		  -Path $excelPath `
		  -WorksheetName $selectedSheet `
	      -StartRow $script:headerRow
		  #-StartRow   $headerRow
		
		if ($data.Count -eq 0) {
            Show-ErrorMessage "No data found in the selected sheet."
            return
        }

        # Convert the date columns in the newly imported data
        if ($global:excelDateCols -and $global:excelDateCols.Count -gt 0) {
            Convert-DateCells -data $data -dateColumns $global:excelDateCols
        }

        # Build hashtable of filter criteria from dynamic panel
        $filters = @{}
        foreach ($ctrl in $panelFilters.Controls) {
            if ($ctrl -is [System.Windows.Forms.TextBox] -and $ctrl.Text) {
                if ($ctrl.Name -match "^tb_(.+)_From$") {
                    $fieldName = $Matches[1]
                    if (-not $filters.ContainsKey($fieldName)) {
                        $filters[$fieldName] = @{}
                    }
                    $filters[$fieldName]["From"] = $ctrl.Text
                }
                elseif ($ctrl.Name -match "^tb_(.+)_To$") {
                    $fieldName = $Matches[1]
                    if (-not $filters.ContainsKey($fieldName)) {
                        $filters[$fieldName] = @{}
                    }
                    $filters[$fieldName]["To"] = $ctrl.Text
                }
                else {
                    if ($ctrl.Name -match "^tb_(.+)$") {
                        $fieldName = $Matches[1]
                        $filters[$fieldName] = $ctrl.Text
                    }
                }
            }
        }

        # Build the dynamic filter scriptblock
        $filterConditions = @()
        foreach ($key in $filters.Keys) {
            $val = $filters[$key]
            if ($val -is [hashtable]) {
                if ($val.ContainsKey("From") -and $val.From) {
                    $filterConditions += "([datetime](`$x.`"" + $key + "`")) -ge ([datetime]'" + $val.From + "')"
                }
                if ($val.ContainsKey("To") -and $val.To) {
                    $filterConditions += "([datetime](`$x.`"" + $key + "`")) -le ([datetime]'" + $val.To + "')"
                }
            }
            else {
                $filterConditions += "([string](`$x.`"" + $key + "`")) -eq '" + $val + "'"
            }
        }

        $filterExpression = $filterConditions -join " -and "
        $code = "param(`$x)`n" + $filterExpression
        $scriptBlock = [ScriptBlock]::Create($code)

        # Filter the data
        $filteredData = $data | ForEach-Object {
            if ($scriptBlock.Invoke($_)) { $_ }
        }

        if ($filteredData.Count -eq 0) {
            Show-ErrorMessage "No rows match these criteria."
            return
        }

        # Random sampling
        $sampleSize = [int]$txtSampleSize.Text
        if ($sampleSize -gt $filteredData.Count) {
            Show-ErrorMessage "Sample size ($sampleSize) exceeds total filtered rows ($($filteredData.Count))."
            return
        }
        $randomSample = $filteredData | Get-Random -Count $sampleSize

        $remainingPool = $filteredData | Where-Object { $randomSample -notcontains $_ }
        $overSampleSize = [int]$txtOverSampleSize.Text
        if ($overSampleSize -gt $remainingPool.Count) {
            Show-ErrorMessage "Oversample size ($overSampleSize) exceeds remaining rows ($($remainingPool.Count))."
            return
        }
        $randomOverSample = $remainingPool | Get-Random -Count $overSampleSize
        
        # Prepare outputs based on selection
        if ($rbGridView.Checked) {
            # Convert dates for GridView display
            $gridViewFilteredData = Convert-DateColumnsForDisplay -dataSet $filteredData -dateColumns $global:excelDateCols -format 'MM/dd/yyyy'
            $gridViewSample = Convert-DateColumnsForDisplay -dataSet $randomSample -dateColumns $global:excelDateCols -format 'MM/dd/yyyy'
            $gridViewOverSample = Convert-DateColumnsForDisplay -dataSet $randomOverSample -dateColumns $global:excelDateCols -format 'MM/dd/yyyy'
            
            # Display results in GridView
            $gridViewFilteredData | Show-GridView -Title 'Filtered results'
            $gridViewSample | Show-GridView -Title 'Random sample'
            $gridViewOverSample | Show-GridView -Title 'Random oversample'
        }

        if ($rbExcel.Checked) {
            # Export results to Excel with proper date formatting
            try {
                $excelOutputPath = $txtExcelPath.Text
                
				$exportResult = Export-ResultsToExcel -filteredData $filteredData -randomSample $randomSample `
					-randomOverSample $randomOverSample -excelPath $excelOutputPath `
					-dateColumns $global:excelDateCols -sampleFormat $txtSamplePrefix.Text
							
                [System.Windows.Forms.MessageBox]::Show($exportResult, "Export Successful", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            catch {
                Show-ErrorMessage "Export Error: $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-LogMessage "Processing error: $($_.Exception)" "ERROR"
		Show-ErrorMessage "Error during processing: $($_.Exception.Message)`nCheck the log for details."
    }
    finally {
        Remove-ProcessingLabel
        foreach ($ctrl in $form.Controls) { $ctrl.Enabled = $true }
    }
})

$btnClear.Add_Click({
    for ($i = 0; $i -lt $clbColumns.Items.Count; $i++) {
        $clbColumns.SetItemChecked($i, $false)
    }
    $clbColumns.Items.Clear()
	$panelFilters.Controls.Clear()
    $txtSampleSize.Text     = ""
    $txtOverSampleSize.Text = ""
    $txtSeed.Text           = ""
    $txtSamplePrefix.Text   = ""
    $rbGridView.Checked     = $false
	$rbExcel.Checked = $true
    $txtExcelPath.Text = Join-Path -Path $scriptDirectory -ChildPath $defaultExcelName
})

# ==========================
# Run Time Handlers
# ==========================

foreach ($control in $form.Controls) {
    if ([string]::IsNullOrEmpty($control.Name)) {
        $null = Ensure-ControlHasName -control $control -prefix "Control_"
    }
    $null = Store-ControlProperties -control $control
}

# Manual resize and scaling hack
$newWidth  = [int]($script:originalFormWidth  * $FormInitHeight)
$newHeight = [int]($script:originalFormHeight * $FormInitwidth)
$form.Width  = $newWidth
$form.Height = $newHeight

$splashForm.Close()
$splashForm.Dispose()
[void]$form.ShowDialog()

Stop-Transcript | Out-Null		
