<%
if request("fname")="" then
%>
   <script>
   function ok(){
            if(document.frm.fname.value==""){
			  document.frm.fname.focus();
			  return false;
			  }
			document.frm.action="OpenTextFile.asp";
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
   set file=fso.OpenTextFile(fname,2, true) ' 1:ForReading 2:ForWriting 8:ForAppending
   response.write(fname&" is created")
   file.close
   set file=nothing
   set fso=nothing
end if
%>