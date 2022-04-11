<%
dim fso, drives, drive, drvname
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
for Each Drive in Drives
    drvname=drive.DriveLetter
    s=drive.DriveType
    select case s
           case 1
	            s="DriveTypeRemovable"
	       case 2
	            s="DriveTypeFixed"
	       case 3
	            s="DriveTypeNetWork"
	       case 4
	            s="DriveTypeCDRom"
	       case 5
	            s="DriveTypeRAMDisk"
	       case else
	            s="DriveTypeUnknown"
end select
response.write(drvname&" DriveDriveType : "&s&"<BR>")
next
set drives=nothing
set fso=nothing
%>