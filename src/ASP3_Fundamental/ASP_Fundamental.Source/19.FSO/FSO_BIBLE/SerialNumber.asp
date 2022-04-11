<%
dim fso, drives, drive, drvname, s
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
for Each Drive in Drives
    If Drive.IsReady Then
       drvname=drive.DriveLetter
       s=hex(drive.SerialNumber)
       response.write(drvname&" Drive SerialNumber : "&s&"<BR>")
	 End If
next
set drives=nothing
set fso=nothing
%>