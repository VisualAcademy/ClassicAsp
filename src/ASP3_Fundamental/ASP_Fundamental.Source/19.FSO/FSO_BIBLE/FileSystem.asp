<%
dim fso, drives, drive, drvname
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
for Each Drive in Drives
    If drive.IsReady Then
       drvname=drive.DriveLetter
       s=drive.FileSystem
       response.write(drvname&" FileSystem : "&s&"<BR>")
   end if
next
set drv=nothing
set fso=nothing
%>