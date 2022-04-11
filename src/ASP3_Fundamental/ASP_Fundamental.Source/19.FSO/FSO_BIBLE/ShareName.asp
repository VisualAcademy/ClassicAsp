<%
dim fso, drives, drive, drvname
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
for Each Drive in Drives
    If Drive.IsReady and Drive.DriveType=3 Then
       drvname=drive.DriveLetter
       s=drive.ShareName
       response.write(drvname&" Drive ShareName : "&s&"<BR>")
	 End If
next
set drives=nothing
set fso=nothing
%>