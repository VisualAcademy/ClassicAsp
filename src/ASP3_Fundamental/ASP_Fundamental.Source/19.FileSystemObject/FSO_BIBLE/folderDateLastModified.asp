<%
if request("fname")="" then
%>
   <script>
       function ok(){
	            if(document.frm.fname.value==""){
				   document.frm.fname.focus();
				   return false;
				   }
				document.frm.action="folderDateLastModified.asp";
				document.frm.submit();
				return true;
	            }
   </script>
   <form name="frm" method="post">
         <input type="text" name="fname"><br>
         <input type="button" value="�Ϸ�" onclick="ok()">
   </form>
<%
else
   fname=request("fname")
   set fso=server.createobject("scripting.filesystemobject")
   set Folder=fso.getFolder(fname)
   If Folder.IsRootFolder Then
      response.write("��Ʈ ������ �ȵ���~")
   Else
      s=folder.DateLastModified
      response.write(fname&" Date Last Modified Is "&s)
   End If
   set Folder=nothing
   set fso=nothing
end if
%>