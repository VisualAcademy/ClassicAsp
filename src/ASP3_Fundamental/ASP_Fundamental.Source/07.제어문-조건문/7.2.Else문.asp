<%
Dim intNumber
intNumber = 10

If intNumber Mod 2 = 1 Then
    Response.Write("홀수입니다.")
Else
	Response.Write("짝수입니다.")
End If
%>