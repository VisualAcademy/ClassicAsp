<%
if (request("fname")="") then
%>
   <script>
   function ok(){
            if(document.frm.FName.value==""){
			  document.frm.FName.focus();
			  return false;
			  }
			document.frm.action="FsoDeleteFile.asp";
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
   file=fso.DeleteFile(FName)
   response.write(FName&" Is Deleted To ")
   set fso=nothing
end if
%>