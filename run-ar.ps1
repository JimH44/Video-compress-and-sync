# Script to run Aeneas for users not familiar with command line
# asks the user to select the video or audio file
# then asks to select the text file
# It will generate the SRT file based on the text file name, 
# then pass the three file names  to the batch file
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = "C:\Users\jenni\Documents\P8CertCourse\arabic"
    Title = 'Select video or audio file'} 
$null = $FileBrowser.ShowDialog()
$vf = $FileBrowser.FileName
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = "C:\Users\jenni\Documents\P8CertCourse\arabic"
    Title = 'Select text file'} 
$null = $FileBrowser.ShowDialog()
$tf = $FileBrowser.FileName
$srt1 = Split-Path -Path $tf -Parent 
$srt2 = ($tf).split('\')[-1] -replace '\.[^.]*$','' 
$srt=$srt1+"\"+$srt2+".srt"
cmd.exe /c "runaen-ar" $vf $tf $srt
