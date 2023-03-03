
param([switch]$Elevated)

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}

'running with full privileges'
'you can mount multiple VHDX, just comma separated list in $PATHS'
$PATHS="C:\ProgramData\Microsoft\Windows\Virtual Hard Disks\CMPFOR.VHDX"
$IMAGES = @()

'mount the VHDXs if not mounted, if mounted, dismount the VHDXs'
foreach ($path in $PATHS) {
$DISK_IMAGE = Get-DiskImage -ImagePath $path
if ($DISK_IMAGE.Attached -eq $False) {Mount-DiskImage  -ImagePath $path -StorageType VHDX -Access ReadWrite}
else {Dismount-DiskImage  -ImagePath $path}
}
'This kills the script process'
stop-process -Id $PID, usefull if run as a shortcut