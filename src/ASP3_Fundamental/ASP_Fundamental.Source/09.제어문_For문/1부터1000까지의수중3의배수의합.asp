<%
	Dim intCount, intSum

	For intCount = 1 To 1000

		If (intCount MOD 3) = 0 Then   '3�ǹ���̸�,,,
			intSum = intSum + intCount   'intSum������ intCount���� ����(1���� 10����)...
		End If

	Next

	Response.Write("1���� 1000������ �� �� 3�� ����� �� : " & intSum)
%>
