<%
Dim a: a = 9
If a Mod 7 = 0 Then
    Response.Write("7�� ���")
ElseIf a Mod 5 = 0 Then
    Response.Write("5�� ���")
ElseIf a Mod 4 = 0 Then
    Response.Write("4�� ���")
Else
    Response.Write("7, 5, 4�� ����� �ƴѼ�")
End If
%>
<script>
var a = 9;
if (a % 7 == 0)
{
    document.write("7�� ���");
}
else if (a % 5 == 0)
{
    document.write("5�� ���");
}
</script>