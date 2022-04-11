<%
if request("fname")="" then
%>
   <script>
   function ok(){
            if(document.frm.fname.value==""){
			  document.frm.fname.focus();
			  return false;
			  }
			document.frm.action="FilesCount.asp";
			document.frm.submit();
            }
   </script>
   <form name="frm">
   <input type="text" name="fname">
   <input type="button" value="Ã£¾Æ!" onclick="ok()">
<%
else
   fname=request("fname")
   set fso=createobject("scripting.filesystemobject")
   set folder=fso.GetFolder(fname)
   set files=folder.files
   s=files.count
   response.write("files are "&s)
   set subfolders=nothing
   set folder=nothing
   set fso=nothing
end if
%>