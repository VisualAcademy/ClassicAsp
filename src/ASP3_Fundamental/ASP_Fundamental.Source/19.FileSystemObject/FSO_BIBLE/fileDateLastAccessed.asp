<%
if request("fname")="" then
%>
   <script>
       function ok(){
	            if(document.frm.fname.value==""){
				   document.frm.fname.focus();
				   return false;
				   }
				document.frm.action="fileDateLastAccessed.asp";
				document.frm.submit();
				return true;
	            }
   </script>
   <form name="frm" method="post">
         <input type="File" name="fname"><br>
         <input type="button" value="�Ϸ�" onclick="ok()">
   </form>
<%
else
   fname=request("fname")
   set fso=server.createobject("scripting.filesystemobject")
   set File=fso.getFile(fname)
   s=file.DateLastAccessed
   response.write(fname&" Date Last Accessed Is "&s)
   set File=nothing
   set fso=nothing
end if
%>