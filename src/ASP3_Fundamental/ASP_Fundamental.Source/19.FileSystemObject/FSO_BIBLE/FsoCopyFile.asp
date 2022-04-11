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
			document.frm.action="FsoCopyFile.asp";
			document.frm.submit();
            }
   </script>
   <form name="frm">
   <input type="file" name="FromFName"><br>
   <input type="Text" name="ToFName"><br>
   <input type="button" value="บนป็!" onclick="ok()">
<%
else
   FromFName=request("FromFName")
   ToFName=request("ToFName")
   set fso=createobject("scripting.filesystemobject")
   file=fso.CopyFile(FromFName,ToFName)
   response.write(FromFName&" Is Copied To "&ToFName)
   set fso=nothing
end if
%>