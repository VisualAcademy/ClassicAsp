<%
if request("fname")="" then
%>
   <script>
   function ok(){
            if(document.frm.fname.value==""){
			  document.frm.fname.focus();
			  return false;
			  }
            if(document.frm.str.value==""){
			  document.frm.str.focus();
			  return false;
			  }
			document.frm.action="writeline.asp";
			document.frm.submit();
            }
   </script>
   <form name="frm">
   <input type="file" name="fname"><br>
   <input type="text" name="str"><br>
   <input type="button" value="¸¸µé¾î!" onclick="ok()">
<%
else
   fname=request("fname")
   str=request("str")
   set fso=createobject("scripting.filesystemobject")
   set file=fso.OpenTextFile(fname,8) ' 1:ForReading 2:ForWriting 8:ForAppending
   file.WriteLine(str)
   response.write(STR&" is writed")
   file.close
   set file=nothing
   set fso=nothing
end if
%>