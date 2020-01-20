#Requires -version 3

Add-Type -AssemblyName System.Windows.Forms

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
}
$null = $FileBrowser.ShowDialog()

If ($FileBrowser.FileNames.Count) {

    $Application = New-Object -ComObject PowerPoint.Application

    ForEach ($FileName in $FileBrowser.FileNames) {

        $Presentations = $Application.Presentations.Open($FileName)
        
        ForEach ($Design in $Presentations.Designs) {
            
            ForEach ($CustomLayout in ($Design.SlideMaster.CustomLayouts | Sort-Object Index -Descending)) {
                Try {
                    $CustomLayout.Delete()
                } Catch {
                    Write-Warning $_.Exception.Message
                }
            }
            
        }
        
    }
    
}

