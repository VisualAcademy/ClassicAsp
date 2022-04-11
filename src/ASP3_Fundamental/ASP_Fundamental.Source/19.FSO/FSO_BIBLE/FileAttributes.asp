<%
if request("fname")="" then
%>
   <script>
       function ok(){
	            if(document.frm.fname.value==""){
				   document.frm.fname.focus();
				   return false;
				   }
				document.frm.action="FileAttributes.asp";
				document.frm.submit();
				return true;
	            }
   </script>
   <form name="frm" method="post">
         <input type="File" name="fname"><br>
         <input type="button" value="¿Ï·á" onclick="ok()">
   </form>
<%
else
   fname=request("fname")
   set fso=server.createobject("scripting.filesystemobject")
   set File=Fso.GetFile(fname)
   s=File.Attributes
   '  0:FileAttrNormal
   '  1:FileAttrReadOnly
   '  2:FileAttrHidden
   '  4:FileAttrSystem
   '  8:FileAttrVolume
   ' 16:FileAttrDirectory
   ' 32:FileAttrArchive
   ' 64:FileAttrAlias
   '128:FileAttrCompressed
   If s=0 Then
      Attr="Normal"
   Else
      If s And 1 Then Attr=Attr&"ReadOnly,"
      If s And 2 Then Attr=Attr&"Hidden,"
      If s And 4 Then Attr=Attr&"System,"
      If s And 8 Then Attr=Attr&"VolumeLabel,"
      If s And 16 Then Attr=Attr&"Directory,"
      If s And 32 Then Attr=Attr&"Archive,"
      If s And 64 Then Attr=Attr&"Alias,"
      If s And 128 Then Attr=Attr&"Compressed"
   End If
   response.write(fname&" Attributes Is "&Attr)
   set file=nothing
   set fso=nothing
end if
%>