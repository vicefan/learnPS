<#
.SYNOPSIS
  Xlsx handling toolkit (skeleton)

.DESCRIPTION
  Import/Export/Append/Update/Merge/Convert helpers for .xlsx files.
#>

[CmdletBinding()]
param()

function Read-Xlsx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [string] $SheetName,
        [switch] $AsDataTable
    )
    if (-not (Test-Path $Path)) { throw "File not found: $Path" }

    Import-Module ImportExcel -ErrorAction Stop
    if ($SheetName) {
        $data = Import-Excel -Path $Path -WorksheetName $SheetName
    } else {
        $data = Import-Excel -Path $Path
    }
    if ($AsDataTable) { return ,$data } else { return $data }
}

function Write-Xlsx {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [object] $Data,
        [string] $Worksheet = 'Sheet1',
        [switch] $AutoSize,
        [switch] $UseCOM  # Excel COM 사용(Excel 설치 필요)
    )

    if ($UseCOM) {
        if (-not (Get-Command -Name 'New-Object' -ErrorAction SilentlyContinue)) { throw "COM not available" }
        if ($PSCmdlet.ShouldProcess($Path, "Write via COM")) {
            $xl = New-Object -ComObject Excel.Application
            $xl.Visible = $false
            $wb = $xl.Workbooks.Add()
            $ws = $wb.Worksheets.Item(1)
            $ws.Name = $Worksheet
            # 간단한 쓰기: 헤더 + 값
            $cols = ($Data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
            for ($c=0; $c -lt $cols.Count; $c++) {
                $ws.Cells.Item(1, $c+1).Value2 = $cols[$c]
            }
            $row = 2
            foreach ($r in $Data) {
                for ($c=0; $c -lt $cols.Count; $c++) {
                    $ws.Cells.Item($row, $c+1).Value2 = $r.$($cols[$c])
                }
                $row++
            }
            $wb.SaveAs((Resolve-Path $Path).Path)
            $wb.Close($false)
            $xl.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    } else {
        Import-Module ImportExcel -ErrorAction Stop
        $params = @{
            Path = $Path
            WorksheetName = $Worksheet
            ClearSheet = $true
            TableName = $Worksheet
        }
        if ($AutoSize) { $params.Add('AutoSize', $true) }
        if ($PSCmdlet.ShouldProcess($Path, "Write via ImportExcel")) {
            $Data | Export-Excel @params
        }
    }
}

function Append-Row {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [psobject] $Row,
        [string] $Worksheet = 'Sheet1'
    )
    Import-Module ImportExcel -ErrorAction Stop
    if ($PSCmdlet.ShouldProcess($Path, "Append row")) {
        $Row | Export-Excel -Path $Path -WorksheetName $Worksheet -Append
    }
}

function Update-Cell {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Worksheet,
        [Parameter(Mandatory)] [int] $Row,
        [Parameter(Mandatory)] [int] $Column,
        [Parameter(Mandatory)] [object] $Value
    )
    # 간단한 방법: Import -> modify in-memory -> write back
    Import-Module ImportExcel -ErrorAction Stop
    $data = Import-Excel -Path $Path -WorksheetName $Worksheet
    if ($Row -le 0 -or $Row -gt $data.Count) { throw "Row out of range" }
    $cols = $data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    if ($Column -le 0 -or $Column -gt $cols.Count) { throw "Column out of range" }
    $colName = $cols[$Column-1]
    $data[$Row-1].$colName = $Value
    Write-Xlsx -Path $Path -Data $data -Worksheet $Worksheet
}

function Get-SheetNames {
    param([Parameter(Mandatory)][string]$Path)
    Import-Module ImportExcel -ErrorAction Stop
    return (Get-ExcelSheetInfo -Path $Path | Select-Object -ExpandProperty Name)
}

<#
Usage examples:
  # Ensure module
  Ensure-Module

  # Read
  $d = Read-Xlsx -Path "C:\temp\sample.xlsx" -SheetName "Sheet1"

  # Write
  Write-Xlsx -Path "C:\temp\out.xlsx" -Data $d -Worksheet "Sheet1" -AutoSize

  # Append
  Append-Row -Path "C:\temp\out.xlsx" -Row ([pscustomobject]@{Name='New'; Value=1})

  # Merge
  Merge-XlsxFiles -Paths "a.xlsx","b.xlsx" -OutPath "merged.xlsx"

  # Convert
  Convert-XlsxToCsv -XlsxPath "a.xlsx" -OutCsvPath "a.csv"

저장 후 필요에 맞게 함수 세부 구현(에러 처리, 대용량 최적화, 스레딩 등)을 추가하세요.
#>