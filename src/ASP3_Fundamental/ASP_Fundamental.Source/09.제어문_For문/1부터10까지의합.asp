<%
	Dim intCount, intSum

	For intCount = 2 To 100 Step 2
	
		intSum = intSum + intCount   'intSum������ intCount���� ����(1���� 10����)...

	Next

	Response.Write("1���� 100������  ¦���� �� : " & intSum)
%>
