<%
dim fso, drives, drive, drvname
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
for Each Drive in Drives
    drvname=drive.DriveLetter
    s=drive.IsReady
    response.write(drvname&" Drive Free Space : "&s&"<BR>")
next
set drives=nothing
set fso=nothing
%>