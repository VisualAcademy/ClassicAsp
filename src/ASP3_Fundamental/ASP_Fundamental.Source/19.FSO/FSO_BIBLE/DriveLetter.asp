<%
dim fso, drives, drive, drvname
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
for Each Drive in Drives
    drvname=drive.DriveLetter
    response.write("��밡�ɵ���̺� : "&drvname&"<BR>")
Next
set drives=nothing
set fso=nothing
%>