<%
' �񱳿����� : =, <>, >, <, >=, <=
' = ������(==) : ������? ������ ��, �׷��� ������ ����
' <> ������(!=) : �ٸ���? �ٸ��� ��, �׷��� ������ ����
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