<%
dim fso, drives, drive, drvname
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
for Each Drive in Drives
    If Drive.IsReady Then
       drvname=drive.DriveLetter
       s=drive.TotalSize/1024
       response.write(drvname&" Drive Total Size : "&s&"<BR>")
	 End If
next
set drives=nothing
set fso=nothing
%>