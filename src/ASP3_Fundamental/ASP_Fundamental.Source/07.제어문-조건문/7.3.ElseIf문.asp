<%
Dim intNumber
intNumber = 7

If intNumber Mod 3 = 0 Then
    Response.Write("3�� ����Դϴ�.")
ElseIf intNumber Mod 4 = 0 Then
	Response.Write("4�� ����Դϴ�.")
Else
	Response.Write("3�� ����� �ƴϰ� 4�� ����� �ƴ�")
End If
%>