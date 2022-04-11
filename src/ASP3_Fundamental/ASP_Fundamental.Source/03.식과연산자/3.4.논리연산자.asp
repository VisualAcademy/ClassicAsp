<%
Option Explicit
Dim a: Dim b
a = 10: b = 5
Response.Write( ((a <> b) And (a = b)) & "<br>" ) '? False
Response.Write( ((a <> b) Or (a = b)) & "<br>" ) '? True
Response.Write( ((a <> b) And _
	Not(a = b)) & "<br>" ) '? True
%>
<SCRIPT LANGUAGE="JavaScript">
var a, b;
a = 10;
b = 5;
document.write( (a > b) && (a == b) , "<br>" ); //? t && f = false
document.write( (a > b) || (a == b) , "<br>" ); //? t || f = true
document.write( !(a > b) ); //? false
</SCRIPT>