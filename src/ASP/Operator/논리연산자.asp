<script>
var a = 3;
var b = 5;

document.write( (a == b) && (a != b) ); // false
document.write( (a == b) || (a != b) ); // true
document.write( !(a == b) ); // false->true
</script>
<%
' 논리연산자 : And(&&), Or(||), Not(!)
' And : 둘 다 참일때에만 참,
' Or : 하나라도 참이면 참, 
' Not : 참이면 거짓으로, 거짓이면 참으로
Dim a: a = 3
Dim b
b = 5
Response.Write( (a = b) And (a <> b) ) ' False
Response.Write( (a = b) Or (a <> b) ) ' True
Response.Write( Not(a = b) ) ' True
%>