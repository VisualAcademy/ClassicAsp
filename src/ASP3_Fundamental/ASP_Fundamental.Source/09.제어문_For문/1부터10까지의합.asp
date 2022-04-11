<%
	Dim intCount, intSum

	For intCount = 2 To 100 Step 2
	
		intSum = intSum + intCount   'intSum변수에 intCount값을 누적(1부터 10까지)...

	Next

	Response.Write("1부터 100까지의  짝수의 합 : " & intSum)
%>
