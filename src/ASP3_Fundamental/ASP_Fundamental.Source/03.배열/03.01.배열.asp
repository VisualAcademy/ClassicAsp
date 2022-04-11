<%
	'[1] 배열선언 : Dim 배열명(첨자수), 첨자수 : 0부터시작 또는 1부터 시작해도 됨.
	Dim strA(3)   'strA(0), strA(1), strA(2), strA(3) --> 총 4개의 공간이 메모리에 할당.

	'[2] 배열 값 대입
	'strA(0) --> Option Base = 1
	strA(1) = "yellow"
	strA(2) = "black"
	strA(3) = "red"

	'[3] 배열값 출력
%>
<%
	'Response.Write strA(1)
%>
<%=strA(1)%><br>
<%=strA(2)%><br>
<%=strA(3)%><br>
<hr>

<table border=1>
<% 	For i = 1 To 3  %>
<tr>
	<td>
		<font color="<%=strA(i)%>"> 여긴 <%=strA(i)%> </font><br>
	</td>
</tr>
<% 	Next %>
</table>