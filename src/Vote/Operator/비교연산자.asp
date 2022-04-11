<%
' 비교연산자 : =, <>, >, <, >=, <=
' = 연산자(==) : 같은지? 같으면 참, 그렇지 않으면 거짓
' <> 연산자(!=) : 다른지? 다르면 참, 그렇지 않으면 거짓
Dim a
Dim b

a = 10
b = 5

Response.Write( (a = b) & "<br />" ) ' False
Response.Write( (a <> b) & "<br />" ) ' True
%>
<script>
var a = 10; var b = 5;
document.write( (a == b) + "<br />") // false
document.write( (a != b) + "<br />") // true
</script>