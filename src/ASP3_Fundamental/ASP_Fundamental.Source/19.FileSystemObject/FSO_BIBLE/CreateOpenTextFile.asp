<%
if request("fname")="" then
%>
   <script>
   function ok(){
            if(document.frm.fname.value==""){
			  document.frm.fname.focus();
			  return false;
			  }
			document.frm.action="CreateOpenTextFile.asp";
			document.frm.submit();
            }
   </script>
   <form name="frm">
   <input type="text" name="fname">
   <input type="button" value="¸¸µé¾î!" onclick="ok()">
<%
else
   fname=request("fname")
   set fso=createobject("scripting.filesystemobject")
   fso.CreateTextFile(fname)
   set file=fso.GetFile(fname)
   set ts=file.OpenAsTextStream(2, true)' 1:ForReading 2:ForWriting 8:ForAppending
   response.write(fname&" is created")
   set ts=nothing
   set file=nothing
   set fso=nothing
end if
%>