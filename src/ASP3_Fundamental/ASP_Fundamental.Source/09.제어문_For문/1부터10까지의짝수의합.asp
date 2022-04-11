<%
	Dim intCount, intSum

	For intCount = 1 To 10 Step 1

		If (intCount MOD 2) = 0 Then   '짝수이면, 
			intSum = intSum + intCount   'intSum변수에 intCount값을 누적(1부터 10까지)...
		End If

	Next

	Response.Write("1부터 10까지의 수 중 짝수의 합 : " & intSum)
%>
