# video-comp-sync.ps1
# of the video-compress-and-sync project
#
# Script to run HandBrakeCLI for users not familiar with command line.
# Asks the user to select the video file to compress
# Then asks for the output folder that syncs with colleagues' machines
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
$programs = ".programs"
$previous_sync_folder_file = $programs\previous_sync.txt
$syncfoldername = "SyncWithColleagues"
$syncfolder = "$here\$syncfoldername"

# Make sure there is a folder to HandBrakeCLI to be in
#
if (-Not (Test-Path ".\$programs")) {
    mkdir ".\$programs"
    if (-Not "$?") {
        $result = [System.windows.forms.messagebox]::show( `
        "I tried to make $here\$programs, but that failed. `
        I need that folder, so I can`'t go any further. `
        Please investigate and fix the problem.")
        exit 1
    }
}

# Check to see if HandBrakeCLI has been installed,
#    and if not, try to install it and remember where it is.
#
$HBname = "Notepad.exe"
$HBname = "README.md"
$HBname = "HandBrakeCLI.exe"
$HBpath = ""

# Look for HandBrakeCLI in current folder\.programs
# or somewhere in the PATH variable
# else install it in $here\$programs
#
if (Test-Path ".\$programs\$HBname") {
    $HBpath = ".\$programs\$HBname"
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
        "$HBname not installed. I'll help you download it and unpack the zip archive into this folder.")

        # The download page lists several products.
        # We want the filename with the latest version number
        #
        $Response = Invoke-WebRequest -URI "https://handbrake.fr/downloads2.php"
        if ($Verbose -match 'HBdownload,') {
            $Response.Links.href
            # $result = [System.windows.forms.messagebox]::show("Response is `"$Response.Links.href`".")
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

        # $downloadURI = Invoke-WebRequest -Uri "https://handbrake.fr/rotation.php?file=$filename"
        # $downloadURI
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
            $Response = Invoke-WebRequest -Uri "$HBurl" -OutFile ".\$programs\$filename"
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
        Expand-Archive -LiteralPath ".\$programs\$filename" -DestinationPath ".\$programs"
    }        
}

# For GUI file selection:
#
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = "$here"
    Title = 'Select video file to compress'} 
$null = $FileBrowser.ShowDialog()
$vf = $FileBrowser.FileName
$result = [System.windows.forms.messagebox]::show("The filename is `"$vf`".")

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

# Now to ask the user.
#
$result = [System.windows.forms.messagebox]::show( `
"Now I'll show you where I think you want the compressed video to go `
so it can synchronize to your colleagues. `
If I got it wrong, please select a different folder.")

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
     InitialDirectory = "$here"
     Title = 'Select folder for compressed video'
     FileName = "$syncfoldername"} 
$null = $FileBrowser.ShowDialog()
$tf = $FileBrowser.FileName

if ("$tf" -ne "$syncfolder") {
    $syncfolder = $tf
}

# Check syncfolder is writable, and remember it for next time.
#
