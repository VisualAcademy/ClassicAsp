<%
if request("fname")="" then
%>
   <script>
       function ok(){
	            if(document.frm.fname.value==""){
				   document.frm.fname.focus();
				   return false;
				   }
				document.frm.action="folderDateLastAccessed.asp";
				document.frm.submit();
				return true;
	            }
   </script>
   <form name="frm" method="post">
         <input type="text" name="fname"><br>
         <input type="button" value="완료" onclick="ok()">
   </form>
<%
else
   fname=request("fname")
   set fso=server.createobject("scripting.filesystemobject")
   set Folder=fso.getFolder(fname)
   If Folder.IsRootFolder Then
      response.write("루트 폴더는 안되지~")
   Else
      s=Folder.DateLastAccessed
      response.write(fname&" Date Last Accessed Is "&s)
   End If
   set Folder=nothing
   set fso=nothing
end if
%>