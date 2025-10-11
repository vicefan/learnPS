function Get-FileName {  
    [CmdletBinding()]  
    Param (   
        [Parameter(Mandatory = $false)]  
        [string]$WindowTitle = 'Open File(s)',

        [Parameter(Mandatory = $false)]
        [string]$InitialDirectory = "$env:USERPROFILE\Documents",

        [Parameter(Mandatory = $false)]
        [string]$Filter,

        [switch]$MultiSelect
    ) 
    Add-Type -AssemblyName System.Windows.Forms

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title  = $WindowTitle
    if ([string]::IsNullOrWhiteSpace($Filter)) {
        $TempFilter = "All Files (*.*)|*.*"
    }
    else {
        $TempFilter = "$Filter (*.$Filter)|*.$Filter"
    }
    $openFileDialog.Filter = $TempFilter
    $openFileDialog.CheckFileExists = $true
    if (![string]::IsNullOrWhiteSpace($InitialDirectory)) { $openFileDialog.InitialDirectory = $InitialDirectory }
    if ($MultiSelect) { $openFileDialog.MultiSelect = $true }

    if ($openFileDialog.ShowDialog().ToString() -eq 'OK') {
        if ($MultiSelect) { 
            $selected = @($openFileDialog.Filenames)
        } 
        else { 
            $selected = $openFileDialog.Filename
        }
    }
    # clean-up
    $openFileDialog.Dispose()

    return $selected
}