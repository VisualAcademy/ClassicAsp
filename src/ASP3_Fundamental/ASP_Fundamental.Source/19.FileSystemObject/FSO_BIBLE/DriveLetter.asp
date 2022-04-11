<%
dim fso, drives, drive, drvname
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
for Each Drive in Drives
    drvname=drive.DriveLetter
    response.write("사용가능드라이브 : "&drvname&"<BR>")
Next
set drives=nothing
set fso=nothing
%>