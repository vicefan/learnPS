Import-Module -Name PSWritePDF
Add-Type -AssemblyName PresentationFramework
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.forms")

# XAML 정의
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="PDF Merger" Height="400" Width="500">
    <DockPanel>
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="5">
            <Button Name="DeleteButton" Height="30" Width="40" Margin="0,0,5,0">Del</Button>
            <Button Name="MergeButton" Height="30" Width="80">Merge</Button>
        </StackPanel>

        <ListBox Name="FileList" AllowDrop="True" Margin="5"
                 Background="#f8f8f8" BorderBrush="#ccc" BorderThickness="1">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <DockPanel>
                        <!-- 인덱스 번호 표시 -->
                        <TextBlock Text="{Binding Index}" Width="25" 
                                   HorizontalAlignment="Left" Padding="4" 
                                   Foreground="#666" FontWeight="Bold"/>

                        <TextBlock Text="{Binding Name}" Padding="4,4,0,4" 
                                   VerticalAlignment="Center"/>

                        <Separator Margin="0,2,0,2" DockPanel.Dock="Bottom"/>
                    </DockPanel>
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>
    </DockPanel>
</Window>
"@

# XAML 로드
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
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
                Merge-PDF -InputFile $ordered -OutputFile $outputPath
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
