<%
Option Explicit
Dim a:	Dim b
a = 5:	b = 8
Response.Write( a Mod b & "<br>") ' 5 / 8 = 0, 5
%>

<SCRIPT LANGUAGE="JavaScript">
var a, b;
a = 5;
b = 8;
document.write( a % b + "<br>" );//? 5 / 8 = ¸ò : 0, ³ª¸ÓÁö : 5
</SCRIPT>