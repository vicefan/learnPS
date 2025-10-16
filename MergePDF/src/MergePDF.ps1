[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Get PSWritePDF module
if (-not (Get-Module -ListAvailable -Name PSWritePDF)) {
    Write-Host "PSWritePDF module is not installed. Installing..." -ForegroundColor Yellow
    Install-Module -Name PSWritePDF -Scope CurrentUser -Force
}

Import-Module -Name PSWritePDF
Add-Type -AssemblyName PresentationFramework
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.forms")

# load XAML
$xamlPath = Join-Path -Path $PSScriptRoot -ChildPath "ui\MergePDF.xaml"
if (-not (Test-Path $xamlPath)) {
    Write-Host "XAML file not found: $xamlPath"
    return
}
# Parse XAML
$xamlString = [System.IO.File]::ReadAllText($xamlPath, [System.Text.Encoding]::UTF8)
$sr = New-Object System.IO.StringReader($xamlString)
$xmlReader = [System.Xml.XmlReader]::Create($sr)
try {
    $window = [Windows.Markup.XamlReader]::Load($xmlReader)
} catch {
    Write-Host "Failed to load XAML: $_"
    return
} finally {
    $xmlReader.Close()
    $sr.Close()
}

if (-not $window) {
    Write-Host "Failed to create Window object."
    return
}

# Get Ctrls
$fileList = $window.FindName("FileList")
$mergeBtn = $window.FindName("MergeButton")
$deleteBtn = $window.FindName("DeleteButton")

if (-not $fileList) { Write-Host "No FileList control found in XAML."; return }
if (-not $mergeBtn) { Write-Host "No MergeButton control found in XAML."; return }
if (-not $deleteBtn) { Write-Host "No DeleteButton control found in XAML."; return }

# --- Update-Idx ---
function Update-Indexes {
    $i = 1
    foreach ($item in $fileList.Items) {
        $item.Index = $i
        $i++
    }
    $fileList.Items.Refresh()
}

# --- Change ---
$draggedItem = $null
$fileList.Add_PreviewMouseLeftButtonDown({
    $draggedItem = $_.OriginalSource.DataContext
    if ($draggedItem -ne $null) {
        [System.Windows.DragDrop]::DoDragDrop($fileList, $draggedItem, 'Move')
    }
})
$fileList.Add_DragOver({ $_.Effects = 'Move'; $_.Handled = $true })
$fileList.Add_Drop({
    param($sender, $e)
    $target = $e.OriginalSource.DataContext
    if ($draggedItem -and $target -and $draggedItem -ne $target) {
        $from = $fileList.Items.IndexOf($draggedItem)
        $to = $fileList.Items.IndexOf($target)
        $fileList.Items.Remove($draggedItem)
        $fileList.Items.Insert($to, $draggedItem)
        Update-Indexes
    }
})

# --- Add Drag&Drop ---
$fileList.Add_Drop({
    param($sender, $e)
    if ($e.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([Windows.DataFormats]::FileDrop)
        foreach ($f in $files) {
            if ($f -like "*.pdf") {
                $obj = [PSCustomObject]@{
                    Index = $fileList.Items.Count + 1
                    Name  = [System.IO.Path]::GetFileName($f)
                    Path  = $f
                }
                $fileList.Items.Add($obj)
            }
        }
    }
})

# --- Delete ---
$deleteBtn.Add_Click({
    $sel = $fileList.SelectedItem
    if ($sel) {
        $fileList.Items.Remove($sel)
        Update-Indexes
    }
})

# --- Merge ---
$mergeBtn.Add_Click({
    $ordered = @()
    foreach ($item in $fileList.Items) {
        $ordered += $item.Path
    }
    if ($ordered.Count -gt 1) {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf"
        $saveFileDialog.Title = "Save Merged PDF"
        $saveFileDialog.FileName = "Merged.pdf"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $outputPath = $saveFileDialog.FileName
            try {
                Merge-PDF -InputFile $ordered -OutputFile $outputPath -IgnoreProtection -WarningAction SilentlyContinue
                [System.Windows.MessageBox]::Show("PDFs merged successfully to `n$outputPath", "Success", 'OK', 'Information')
            } catch {
                [System.Windows.MessageBox]::Show("Error merging PDFs: $_", "Error", 'OK', 'Error')
            }
        }
    } else {
        [System.Windows.MessageBox]::Show("Please add at least two PDF files to merge.", "Warning", 'OK', 'Warning')
    }
})

# --- Execute ---
$window.ShowDialog() | Out-Null