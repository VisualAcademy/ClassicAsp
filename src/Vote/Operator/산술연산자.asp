<%
' ��������� : +, -, *, /, Mod
'[1] ���� ����
Dim a
Dim b
Dim c
'[2] ���� �ʱ�ȭ
a = 10
b = 5
c = b Mod a ' b Mod a -> b / a : �� : 0, ������ : 5
'[3] ���� ���
Response.Write(c & "<br />") ' 5
'[4] ������ ���� ���
Response.Write(TypeName(c) & "<br />") ' Integer    
%>
<script>
var a = 10; var b = 5; var c = b % a;
window.document.write(c + "<br />");
document.write(typeof(c) + "<br />"); // number
</script>