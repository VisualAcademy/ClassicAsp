<script>
var a = 3;
var b = 5;

document.write( (a == b) && (a != b) ); // false
document.write( (a == b) || (a != b) ); // true
document.write( !(a == b) ); // false->true
</script>
<%
' �������� : And(&&), Or(||), Not(!)
' And : �� �� ���϶����� ��,
' Or : �ϳ��� ���̸� ��, 
' Not : ���̸� ��������, �����̸� ������
Dim a: a = 3
Dim b
b = 5
Response.Write( (a = b) And (a <> b) ) ' False
Response.Write( (a = b) Or (a <> b) ) ' True
Response.Write( Not(a = b) ) ' True
%>