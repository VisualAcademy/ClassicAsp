<%
if ((request("Fromfname")="") or (request("Tofname")=""))then
%>
   <script>
   function ok(){
            if(document.frm.FromFName.value==""){
			  document.frm.FromFName.focus();
			  return false;
			  }
            if(document.frm.ToFName.value==""){
			  document.frm.ToFName.focus();
			  return false;
			  }
			document.frm.action="FileMove.asp";
			document.frm.submit();
            }
   </script>
   <form name="frm">
   <input type="file" name="FromFName"><br>
   <input type="Text" name="ToFName"><br>
   <input type="button" value="¿Å°Ü!" onclick="ok()">
<%
else
   FromFName=request("FromFName")
   ToFName=request("ToFName")
   set fso=createobject("scripting.filesystemobject")
   set file=fso.GetFile(FromFName)
   file.move(ToFName)
   response.write(FromFName&" Is Moved To "&ToFName)
   set file=nothing
   set fso=nothing
end if
%>