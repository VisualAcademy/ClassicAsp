<%
   set fso=createobject("scripting.filesystemobject")
   set folder=fso.GetSpecialFolder
   fname=folder.name
   response.write("system folder is? "&fname)
   set folder=nothing
   set fso=nothing
%>