<%
if request("fromfname")="" then
%>
   <script>
   function ok(){
            if(document.frm.fromfname.value==""){
			  document.frm.fromfname.focus();
			  return false;
			  }
            if(document.frm.tofname.value==""){
			  document.frm.tofname.focus();
			  return false;
			  }
			document.frm.action="MoveFolder.asp";
			document.frm.submit();
            }
   </script>
   <form name="frm">
   <input type="text" name="fromfname">
   <input type="text" name="tofname">
   <input type="button" value="옮겨!" onclick="ok()">
<%
else
   fromfname=request("fromfname")
   tofname=request("tofname")
   set fso=createobject("scripting.filesystemobject")
   set folder=fso.getfolder(fromfname)
   folder.Copy(tofname)
   set folder=nothing
   set fso=nothing
   %><a href="MainMenu.asp"><%=fromfname%>을 <%=tofname%>으로 복사했습니다.</a><%
end if
%>