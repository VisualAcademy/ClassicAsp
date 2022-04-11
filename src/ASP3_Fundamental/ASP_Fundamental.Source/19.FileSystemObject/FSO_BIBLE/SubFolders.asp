<%
if request("fname")="" then
%>
   <script>
       function ok(){
	            if(document.frm.fname.value==""){
				   document.frm.fname.focus();
				   return false;
				   }
				document.frm.action="SubFolders.asp";
				document.frm.submit();
				return true;
	            }
   </script>
   <form name="frm" method="post">
         <input type="text" name="fname"><br>
         <input type="button" value="¿Ï·á" onclick="ok()">
   </form>
<%
else
   fname=request("fname")
   set fso=server.createobject("scripting.filesystemobject")
   set Folder=fso.getfolder(fname)
   set SubFolders=folder.SubFolders
   for each SubFolder in SubFolders
       SubFolderName=SubFolder.Name
	   response.write(SubFolderName&"<br>")
   next
   set SubFolders=nothing
   set Folder=nothing
   set fso=nothing
end if
%>