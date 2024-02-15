# video-comp-sync.ps1
# of the video-compress-and-sync project
#
# Script to run HandBrakeCLI for users not familiar with command line.
# Depends on someone setting Resilio Sync between folder SyncWithOthers
# and other machine or machines.
# If the variable $LetUserChangeSyncFolder is true,
#    asks for the output folder that syncs with colleagues' machines
#    using GUI folder chooser
# Asks the user to select the video file to compress
#    using GUI file chooser
# It will compress the video using built-in parameters, 
# with output to the syncing folder.
#
# To help with debugging:
# $Verbose = "HB,HBdownload,extract,"
$Verbose = ""

# Set-ExecutionPolicy unrestricted

# Make sure we can do dialogs.
#
Add-Type -AssemblyName System.Windows.Forms

# Assuming this script is in the folder where the videos are,
# we'll get the current working directory.
#
$here = pwd

# some other assumptions:
$LetUserChangeSyncFolder = $false
$LetUserChangeSyncFolder = $true

$support = ".support"     # folder to put HandbrakeCLI etc

# for remembering previous selections
#
$previous_sync_folder_file = "$support\previous_sync.txt"
$syncfoldername = "SyncWithOthers"
$syncfolder = "$here\$syncfoldername"

# Make sure there is a folder for HandBrakeCLI to be in
#
if (-Not (Test-Path ".\$support")) {
    mkdir ".\$support"
    if (-Not "$?") {
        $result = [System.windows.forms.messagebox]::show( `
        "I tried to make $here\$support, but that failed. `
        I need that folder, so I can`'t go any further. `
        Please investigate and fix the problem.")
        exit 1
    }
}

# Check to see if HandBrakeCLI has been installed,
#    and if not, try to install it and remember where it is.
#
# $HBname = "Notepad.exe"
# $HBname = "README.md"
$HBname = "HandBrakeCLI.exe"
$HBpath = ""

# Look for HandBrakeCLI in current folder
# or current folder\$support
# or somewhere in the PATH variable
# else install it in $here\$support
#
if (Test-Path ".\$HBname") {
    $HBpath = ".\$HBname"
    if ($Verbose -match 'HB,') {
        $result = [System.windows.forms.messagebox]::show("The path to $HBname is `"$HBpath`".")
    }
}
elseif (Test-Path ".\$support\$HBname") {
    $HBpath = ".\$support\$HBname"
    if ($Verbose -match 'HB,') {
        $result = [System.windows.forms.messagebox]::show("The path to $HBname is `"$HBpath`".")
    }
}
else {
    $HBpath = (Get-Command "$HBname" -ErrorAction SilentlyContinue | Select-Object Source).Source
    
    if ($null -ne $HBpath) {  
        if ($Verbose -match 'HB,') {
            $result = [System.windows.forms.messagebox]::show("The path to $HBname is `"$HBpath`".")            
        }
    } else {
        $result = [System.windows.forms.messagebox]::show( `
        "$HBname not installed. I'll help you download it and unpack the zip archive into this folder.")

        # The download page lists several products.
        # We want the filename with the latest version number
        #
        $Response = Invoke-WebRequest -URI "https://handbrake.fr/downloads2.php"
        if ($Verbose -match 'HBdownload,') {
            $Response.Links.href
        }
        $filename  = [regex]::match($Response.Links.href,'file=(HandBrakeCLI-[\d\.]+-win-x86_64.zip)').Groups[1].Value
        if ($filename) {
            $result = [System.windows.forms.messagebox]::show("Filename to get is `"$filename`".")
        } else {
            $result = [System.windows.forms.messagebox]::show( `
                "Sorry, I wasn't able to work out which file to get. `
                Please use a web browser and go to `"https://handbrake.fr/downloads2.php`" `
                and get the installer for HandBrakeCLI for Windows and unpack it in this folder.")
            exit 2
        }

        if ($Verbose -match 'HBdownload,') {
            $result = [System.windows.forms.messagebox]::show( `
                "The intermediate URL is https://handbrake.fr/rotation.php?file=$filename `
                I'll get it now.")
        }

        $HBversion  = [regex]::match($filename,'HandBrakeCLI-([\d\.]+)-win-x86_64.zip').Groups[1].Value
        if ($Verbose -match 'HBdownload,') {
            $HBversion
        }

        $HBurl = "https://github.com/HandBrake/HandBrake/releases/download/$HBversion/$filename"
        $HBversion  = [regex]::match($filename,'HandBrakeCLI-([\d\.]+)-win-x86_64.zip').Groups[1].Value
        if ($Verbose -match 'HBdownload,') {
            $HBurl
        }

        # Turn off the download progress bar to save 90% of download time
        #
        $ProgressPreference = 'SilentlyContinue'
        
        # Now to download the file
        #
        try {
            $Response = Invoke-WebRequest -Uri "$HBurl" -OutFile ".\$support\$filename"
            # This will only execute if the Invoke-WebRequest is successful.
            $StatusCode = $Response.StatusCode
        } catch {
            $StatusCode = $_.Exception.Response.StatusCode.value__
        }
        $HBversion  = [regex]::match($filename,'HandBrakeCLI-([\d\.]+)-win-x86_64.zip').Groups[1].Value
        if ($Verbose -match 'HBdownload,') {
            $StatusCode
        }

        if ($StatusCode) {
            $result = [System.windows.forms.messagebox]::show( `
            "Sorry, I was not able to download `"$filename`". `
            Please use a web browser to go to to `"https://handbrake.fr/downloads2.php`" `
            and get the installer for HandBrakeCLI for Windows and unpack it in this folder.")
            exit 3
        }

        # Zip archive downloaded -- now to unpack it here.
        #
        $result = [System.windows.forms.messagebox]::show( `
        "I have downloaded `"$filename`" -- now I'll unpack the archive here.")
        Expand-Archive -LiteralPath ".\$support\$filename" -DestinationPath ".\$support"
    }        
}

# Find out where to send the compressed file
#
# If this has been selected before,
#    get the previous value.
#
if ((Test-Path "$previous_sync_folder_file")) {
    if($file = Get-Content $previous_sync_folder_file 2>$null) {
        if(-Not $syncfolder -eq $file) {
            $syncfolder = $file
        }
    }
}

if ($LetUserChangeSyncFolder) {

    # Now to ask the user where the Sync folder is.
    #
    $result = [System.windows.forms.messagebox]::show( `
    "Now I'll show you where I think you want the compressed video to go `
    so it can synchronize to your colleagues. Please scroll down to see it. `
    If I got it wrong, please select a different folder.")
    
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        RootFolder = 'Desktop'
        ShowNewFolderButton = $false
        SelectedPath = "$syncfolder"
        Description = 'Select folder to put compressed video'
    }
    $result = $dialog.ShowDialog()
    # If the user cancelled, leave $syncfolder unchanged.
    #
    if ("$result" -eq [System.Windows.Forms.DialogResult]::Cancel) {
        $result = [System.windows.forms.messagebox]::show( `
        "OK, I`'ll leave the folder for compressed videos `
        as `"$syncfolder`".")
    }
    else {$sf = $dialog.SelectedPath}

    # If the user changes the sync directory,
    #    remember that for next  time.
    #
    if ("$sf" -ne "$syncfolder") {
        $syncfolder = $sf

        # Check syncfolder is writable, and remember it for next time.
        #
        Try { [io.file]::OpenWrite("$syncfolder\thing.txt").close() }
        Catch { Write-Warning "Unable to write to $outsyncfolder"
            exit 6
        }

        try {
            "$syncfolder" > $previous_sync_folder_file
        }
        catch {
            Write-Warning "Unable to write to file $previous_sync_folder_file"
            exit 7
        }
    }
} 

# For GUI file selection of the video to compress:
#
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = "$here"
    Title = 'Select video file to compress'} 
$result = $FileBrowser.ShowDialog()
$video_file_path = $FileBrowser.FileName

if ("$result" -eq [System.Windows.Forms.DialogResult]::Cancel) {
    $result = [System.windows.forms.messagebox]::show( `
    "No video selected. Goodbye.")
    exit 10
} else {"$result"}

$result = [System.windows.forms.messagebox]::show("The filename is `"$video_file_path`".")
$video_file_name = Split-Path $video_file_path -leaf

# Now to compress the video
#
$result = [System.windows.forms.messagebox]::show("I'm ready to compress the video now. `
    Please leave the computer running until the blue window goes away, `
    and then also leave it on and connected to the internet so the compressed video `
    can sync to your colleagues.")
try {
    &$HBpath -e x264 -q 28 -r 15 -B 64 -X 1280 -O with -i "$video_file_path" `
        -o "$syncfolder\$video_file_name"
}
catch {
    Write-Warning "Unable to compress video to "$syncfolder\$video_file_name""
    exit 9
}
