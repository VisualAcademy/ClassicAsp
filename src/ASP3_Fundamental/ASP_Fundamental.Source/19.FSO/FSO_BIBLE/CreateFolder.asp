<%
if request("fname")="" then
%>
   <script>
   function ok(){
            if(document.frm.fname.value==""){
			  document.frm.fname.focus();
			  return false;
			  }
			document.frm.action="CreateFolder.asp";
			document.frm.submit();
            }
   </script>
   <form name="frm">
   <input type="text" name="fname">
   <input type="button" value="�����!" onclick="ok()">
<%
else
   fname=request("fname")
   set fso=createobject("scripting.filesystemobject")
   fso.createfolder(fname)
   set fso=nothing
   %><a href="nainmenu.asp"><%=fname%>�� �����߽��ϴ�.</a><%
end if
%>