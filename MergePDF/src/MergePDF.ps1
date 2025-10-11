# PSWritePDF 모듈 확인 및 설치
if (-not (Get-Module -ListAvailable -Name PSWritePDF)) {
    Write-Host "PSWritePDF module is not installed. Installing..." -ForegroundColor Yellow
    Install-Module -Name PSWritePDF -Scope CurrentUser -Force
}

Import-Module -Name PSWritePDF
Add-Type -AssemblyName PresentationFramework
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.forms")

# XAML 로드 (외부 파일)
$xamlPath = Join-Path -Path $PSScriptRoot -ChildPath "ui\MergePDF.xaml"
# 파일을 원문으로 읽고 XML로 파싱 (인코딩 문제 방지)
$xamlString = [System.IO.File]::ReadAllText($xamlPath)
[xml]$xamlXml = $xamlString
$reader = New-Object System.Xml.XmlNodeReader $xamlXml
$window = [Windows.Markup.XamlReader]::Load($reader)

# 창 위치
$window.WindowStartupLocation = 'CenterScreen'

# 컨트롤
$fileList = $window.FindName("FileList")
$mergeBtn = $window.FindName("MergeButton")
$deleteBtn = $window.FindName("DeleteButton")

# --- 인덱스 갱신 함수 ---
function Update-Indexes {
    $i = 1
    foreach ($item in $fileList.Items) {
        $item.Index = $i
        $i++
    }
    $fileList.Items.Refresh()
}

# --- Drag 순서 변경 ---
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

# --- 파일 Drag & Drop 추가 ---
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

# --- 삭제 버튼 ---
$deleteBtn.Add_Click({
    $sel = $fileList.SelectedItem
    if ($sel) {
        $fileList.Items.Remove($sel)
        Update-Indexes
    }
})

# --- 병합 버튼 ---
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

# --- 실행 ---
$window.ShowDialog() | Out-Null