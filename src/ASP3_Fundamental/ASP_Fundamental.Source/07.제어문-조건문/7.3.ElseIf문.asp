<%
Dim intNumber
intNumber = 7

If intNumber Mod 3 = 0 Then
    Response.Write("3의 배수입니다.")
ElseIf intNumber Mod 4 = 0 Then
	Response.Write("4의 배수입니다.")
Else
	Response.Write("3의 배수도 아니고 4의 배수도 아님")
End If
%>