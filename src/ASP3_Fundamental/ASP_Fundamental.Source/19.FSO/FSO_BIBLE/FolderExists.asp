<%
if request("fname")="" then
%>
   <script>
   function ok(){
            if(document.frm.fname.value==""){
			  document.frm.fname.focus();
			  return false;
			  }
			document.frm.action="FolderExists.asp";
			document.frm.submit();
            }
   </script>
   <form name="frm">
   <input type="text" name="fname">
   <input type="button" value="찾아!" onclick="ok()">
<%
else
   fname=request("fname")
   set fso=createobject("scripting.filesystemobject")
   s=fso.FolderExists(fname)
   set fso=nothing
   %><a href="MainMenu.asp"><%=fname%>가 있냐? ==> <%=s%></a><%
end if
%>