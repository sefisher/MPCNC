################################################################################################
#
# This PowerShell script takes CNC code exported from Easel and converts it for use in a MPCNC 
#
# Author: Steve Fisher
# Date:   6 Dec 2020
#
# ***You are free to use and modify - but do so at your own risk.***
#
# Purpose: This modifies an Easel GCODE file to prepare it for use in an MARLIN based MPCNC machine
# that is ready with the bit zeroed on the workpiece (X=0, Y=0, Z=[height of z-probe plate]). My 
# flow is: 1) Export from Easel, 2) Run this script, 3) Upload file to OctoPi server, 4) Zero bit 
# on the workpiece (0,0,z-probe-height), 5) Run GCODE, 6) Turn on spindle when paused at safeheight,
# 7) Click LCD button to continue cutting.
#
# What the script does: This script prompts you to open an NC file exported by Easel (usually "Untiled.nc"
# in the Downloads folder). Then it confirms the file units are set to mm. It asks the user for a 
# "safe height" to start. Then it adds a short set of GCODEs to the file's beginning to:
#    1) turn on absolute positioning, 
#    2) set the bit's current position to {0,0,[height of z-probe plate defined below as $zProbeOffset]}
#    3) raises the bit the safeheight defined below as $defaultSafeHeight or input by the user
#    4) pauses the machine waiting for an LCD button press at the front panel (to allow turning 
#       on the spindle and getting it the right speed).
#
# NOTE: For most Windows machines you have to run the script bypassing file security policy to allow it 
#       to open and modify the GCODE file.  To do that:
#       
#       -Create a shortcut to the script on your desktop (or somewhere you can find it).
#       -Right-click the shortcut and click Properties.
#       -Click the Shortcut tab.
#       -Edit the "Target" field to add "powershell.exe -ExecutionPolicy Bypass -File" in front of the file
#        (e.g.,  "powershell.exe -ExecutionPolicy Bypass -File D:\3D Objects\CNC Parts\Easel to CNC script.ps1")
#
################################################################################################

#some parameters
$defaultSafeHeight = 5.000  #mm above Z=0
$defaultMoveSpeed = 3000    # mm/min for repositioning bit when not cutting
$firstLine = ";MODIFIED FOR STEVE FISHER'S MPCNC MACHINE USING EASEL - to - MPCNC SCRIPT"
$Downloads = "D:\Downloads\"
$SaveFolder = "Q:\"
$zProbeOffset = 0.75  #this should be the thickness of the metal plate used to zero your bit height on top of the work piece

Write-Host "`r`n`r`n--Running EASEL to MPCNC--`r`n`r`n"

#get Easel output file
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = $Downloads; Filter = 'NC Files (*.nc)|*.nc'}
if($FileBrowser.ShowDialog() -eq 'Ok'){

    $lines = Get-Content $FileBrowser.FileName
    $numLines = $lines.Count

    #check for metric - expect the first line to set inches or mm (G20 or G21) when exported from Easel
    if($lines[0] -eq "G21") {
        Write-Host "Verfified NC file units is set for metric (mm).`r`n`r`n"

        #get the safe height
        if (!($safeheight = Read-Host "What is the Project Safe Height in mm (default = $defaultSafeHeight mm)?")) {$safeheight = $defaultSafeHeight}
        #get the move speed
        if (!($movespeed = Read-Host "`r`nWhat speed in mm/min do you want to move when not cutting (default = $defaultMoveSpeed mm/min)?")) {$movespeed = $defaultMoveSpeed}

#Insert opening GCODE:
$openingText = 
@"
$firstLine
;To use this position to the zero starting point for piece using the Zprobe.
;then it will pause
;remove the probe & turn on the spindle
;then hit ok on the LCD
G21 ; Set to metric (mm)
G90 ; Absolute positioning, just in case
G92 X0 Y0 Z0 ; Set Current position to 0, all axes
G92 Z$zProbeOffset ; Account for probe thickness (set your thickness) - this is a coordinate system offset
G00 Z$safeheight F500 ; Raise Z probe off off of surface
M00 ; pause for LCD button press
"@

        #open dilog box to save new file
        $FileSave = New-Object System.Windows.Forms.SaveFileDialog -Property @{InitialDirectory = $SaveFolder; Filter = "GCODE Files (*.gcode)|*.gcode"}

        if($FileSave.ShowDialog() -eq 'Ok'){

            $newFilename = $FileSave.FileName
        
            Write-Host "`r`nSaving the new file to: $newFilename."
        
            $newfile = New-Item $FileSave.FileName -type "file" -Value $openingText -Force

            #process the Easel file to find where G0 codes are (these are moves done at the safe height set in Easel) and set the new move speed
            for($i = 0; $i -lt $numLines; $i++){ 
                if($lines[$i].StartsWith("G0 ")){
                    $old = $lines[$i]
                    $lines[$i] = "$old F$defaultMoveSpeed" 
                    #Write-Host $lines[$i]    
                }
            }

            #save the file
            Add-Content $FileSave.FileName $lines
    
        } else {
            Write-Host "`r`n`r`n--Cancelled--"
        }

    }elseif ($lines[0] -eq "G20"){ # G20 is code for inches - terminate
        Write-Host "`r`n`r`nERROR - WRONG DIMENSIONS USED! MUST SET DIMENSIONS TO mm IN EASEL."
    }elseif ($lines[0] -eq $firstLine){ # file already modified - terminate
        Write-Host "`r`n`r`nERROR - YOU ARE TRYING TO PROCESS A FILE THAT HAS ALREADY BEEN MODIFIED BY THIS SCRIPT."
    }else {
        Write-Host "`r`n`r`nERROR - CAN'T DETERMINE NC FILE UNITS. FIRST LINE SHOULD BE G21 (mm) OR G20 (inches)." 
    }

}else{
    Write-Host "`r`n`r`n--Cancelled--`r`n`r`n"
}

Read-Host -Prompt "`r`n`r`n::Press Enter to exit:"