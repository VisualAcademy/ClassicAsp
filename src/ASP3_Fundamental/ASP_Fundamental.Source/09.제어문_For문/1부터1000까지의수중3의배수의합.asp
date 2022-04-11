<%
	Dim intCount, intSum

	For intCount = 1 To 1000

		If (intCount MOD 3) = 0 Then   '3의배수이면,,,
			intSum = intSum + intCount   'intSum변수에 intCount값을 누적(1부터 10까지)...
		End If

	Next

	Response.Write("1부터 1000까지의 수 중 3의 배수의 합 : " & intSum)
%>
