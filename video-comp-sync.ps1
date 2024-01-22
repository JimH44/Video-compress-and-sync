# video-comp-sync.ps1
# of the video-compress-and-sync project
#
# Script to run HandBrakeCLI for users not familiar with command line
# asks the user to select the video or audio file to compress
# Then asks for the output folder that syncs with colleagues' machines
# It will compress the video using built-in parameters, 
# with output to the syncing folder
#
# Assuming this script is in the folder where the videos are,
# we'll get the current working directory.
#
$here = pwd

# To help with debugging:
# $Verbose = "HB,""
$Verbose = ""

# Check to see if HandBrakeCLI has been installed,
#    and if not, try to install it and remember where it is.
#
$HBname = "Notepad.exe"
$HBname = "README.md"
$HBname = "HandBrakeCLI.exe"
$HBpath = ""

if (Test-Path ".\$HBname") {
    $HBpath = ".\$HBname"
    if ($Verbose -match 'HB,') {
        $result = [System.windows.forms.messagebox]::show("The path to $HBname is `"$HBpath`".")
    }
}
else {
    $HBpath = (Get-Command "$HBname" -ErrorAction SilentlyContinue | Select-Object Source).Source
    # $result = [System.windows.forms.messagebox]::show("After test, the path to $HBname is `"$HBpath`".")
    
    if ($HBpath -ne $null ) {  
        if ($Verbose -match 'HB,') {
            $result = [System.windows.forms.messagebox]::show("The path to $HBname is `"$HBpath`".")            
        }
    } else {
        $result = [System.windows.forms.messagebox]::show( `
        "$HBname not installed. Please download it and unpack the zip archive into this folder.")
        exit 1
    }        
}
exit
# For GUI file selection:
#
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = "$here"
    Title = 'Select video file to compress'} 
$null = $FileBrowser.ShowDialog()
$vf = $FileBrowser.FileName
$result = [System.windows.forms.messagebox]::show("The filename is `"$vf`".")
exit
# Exit-PSSession
# $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
#     InitialDirectory = "C:\Users\jenni\Documents\P8CertCourse\arabic"
#     Title = 'Select text file'} 
# $null = $FileBrowser.ShowDialog()
# $tf = $FileBrowser.FileName
# $srt1 = Split-Path -Path $tf -Parent 
# $srt2 = ($tf).split('\')[-1] -replace '\.[^.]*$','' 
# $srt=$srt1+"\"+$srt2+".srt"
# cmd.exe /c "runaen-ar" $vf $tf $srt
