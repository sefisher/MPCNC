
 This PowerShell script takes CNC code exported from Easel and converts it for use in a MPCNC 

 Author: Steve Fisher
 Date:   6 Dec 2020

 ***You are free to use and modify - but do so at your own risk.***

 Purpose: This modifies an Easel GCODE file to prepare it for use in an MARLIN based MPCNC machine
 that is ready with the bit zeroed on the workpiece (X=0, Y=0, Z=[height of z-probe plate]). My 
 flow is: 1) Export from Easel, 2) Run this script, 3) Upload file to OctoPi server, 4) Zero bit 
 on the workpiece (0,0,z-probe-height), 5) Run GCODE, 6) Turn on spindle when paused at safeheight,
 7) Click LCD button to continue cutting.

 What the script does: This script prompts you to open an NC file exported by Easel (usually "Untiled.nc"
 in the Downloads folder). Then it confirms the file units are set to mm. It asks the user for a 
 "safe height" to start. Then it adds a short set of GCODEs to the file's beginning to:
    1) turn on absolute positioning, 
    2) set the bit's current position to {0,0,[height of z-probe plate defined below as $zProbeOffset]}
    3) raises the bit the safeheight defined below as $defaultSafeHeight or input by the user
    4) pauses the machine waiting for an LCD button press at the front panel (to allow turning 
       on the spindle and getting it the right speed).

 NOTE: For most Windows machines you have to run the script bypassing file security policy to allow it 
       to open and modify the GCODE file.  To do that:
       
       -Create a shortcut to the script on your desktop (or somewhere you can find it).
       -Right-click the shortcut and click Properties.
       -Click the Shortcut tab.
       -Edit the "Target" field to add "powershell.exe -ExecutionPolicy Bypass -File" in front of the file
        (e.g.,  "powershell.exe -ExecutionPolicy Bypass -File D:\3D Objects\CNC Parts\Easel to CNC script.ps1")

