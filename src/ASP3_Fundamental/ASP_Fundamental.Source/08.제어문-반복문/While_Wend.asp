<%
	'While문으로 1~100까지의 수 중 짝수의 합을 구하는 프로그램
	Dim Sum, Count

	Sum = 0
	Count = 1

	While Count <= 100
		If Count Mod 2 = 0 Then
			Sum = Sum + Count
		End If
		Count = Count + 1   '카운터 변수 증가
	Wend

	Response.Write("<center>1부터 100까지의 짝수의 합은 " & Sum & " 입니다</center>")
%>