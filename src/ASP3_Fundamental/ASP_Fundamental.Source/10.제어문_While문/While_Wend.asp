<%
	'While������ 1~100������ �� �� ¦���� ���� ���ϴ� ���α׷�
	Dim Sum, Count

	Sum = 0
	Count = 1

	While Count <= 100
		If Count Mod 2 = 0 Then
			Sum = Sum + Count
		End If
		Count = Count + 1   'ī���� ���� ����
	Wend

	Response.Write("<center>1���� 100������ ¦���� ���� " & Sum & " �Դϴ�</center>")
%>