<%
	Dim intCount, intSum

	For intCount = 1 To 10 Step 1

		If (intCount MOD 2) = 0 Then   '¦���̸�, 
			intSum = intSum + intCount   'intSum������ intCount���� ����(1���� 10����)...
		End If

	Next

	Response.Write("1���� 10������ �� �� ¦���� �� : " & intSum)
%>
