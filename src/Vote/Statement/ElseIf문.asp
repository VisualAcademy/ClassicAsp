<%
Dim a: a = 9
If a Mod 7 = 0 Then
    Response.Write("7의 배수")
ElseIf a Mod 5 = 0 Then
    Response.Write("5의 배수")
ElseIf a Mod 4 = 0 Then
    Response.Write("4의 배수")
Else
    Response.Write("7, 5, 4의 배수가 아닌수")
End If
%>
<script>
var a = 9;
if (a % 7 == 0)
{
    document.write("7의 배수");
}
else if (a % 5 == 0)
{
    document.write("5의 배수");
}
</script>