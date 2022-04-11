<%
dim fso, drives, drive, drvname, s
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
for Each Drive in Drives
	If Drive.IsReady Then
       drvname=drive.DriveLetter
       s=drive.AvailableSpace/1024
       response.write(drvname&" 사용가능용량 : "&s&" KB<br>")
   End If
Next
set drives=nothing
set fso=nothing
%>