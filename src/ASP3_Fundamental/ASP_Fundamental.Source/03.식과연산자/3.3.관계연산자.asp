<%
' 관계(비교) 연산자 : >, >=, <, <=, =, <>
Dim a: Dim b
a = 5
b = 10
Response.Write( a  =  b ) ' = :  a와 b가 같은지?
Response.Write("<br>")
Response.Write( a <>  b ) ' <> : a와 b가 다른지?
%><br>
<SCRIPT LANGUAGE="JavaScript">
var a, b;
a = 5;
b = 10;
document.write( a == b , "<br>" ); //같으냐?
document.write( a != b );//다르냐?
</SCRIPT>