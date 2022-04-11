<%
dim fso, drive
dim drives(3)
drives(0)="a:"
drives(1)="b:"
drives(2)="c:"
drives(3)="d:"
Set fso=createobject("scripting.filesystemobject")
for each drive in drives
    s=fso.DriveExists(drive)
    response.write(drive&" Drive Exist : "&s&"<BR>")
next
set fso=nothing
%>