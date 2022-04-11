<%
if (request("fname")="")then
%>
   <script>
   function ok(){
            if(document.frm.FName.value==""){
			  document.frm.FName.focus();
			  return false;
			  }
			document.frm.action="FileDelete.asp";
			document.frm.submit();
            }
   </script>
   <form name="frm">
   <input type="file" name="FName"><br>
   <input type="button" value="»èÁ¦!" onclick="ok()">
<%
else
   FName=request("FName")
   set fso=createobject("scripting.filesystemobject")
   set file=fso.GetFile(FName)
   file.Delete
   response.write(FName&" Is Deleted")
   set file=nothing
   set fso=nothing
end if
%>