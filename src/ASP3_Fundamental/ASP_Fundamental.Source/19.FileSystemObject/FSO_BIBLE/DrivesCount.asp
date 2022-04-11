<%
dim fso, drives, s
Set fso=createobject("scripting.filesystemobject")
set drives=fso.drives
s=drives.count

response.write("Total Drive Count : "&s&" EA")
set drives=nothing
set fso=nothing
%>