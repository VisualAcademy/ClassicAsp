<%
	'[1] �迭���� : Dim �迭��(÷�ڼ�), ÷�ڼ� : 0���ͽ��� �Ǵ� 1���� �����ص� ��.
	Dim strA(3)   'strA(0), strA(1), strA(2), strA(3) --> �� 4���� ������ �޸𸮿� �Ҵ�.

	'[2] �迭 �� ����
	'strA(0) --> Option Base = 1
	strA(1) = "yellow"
	strA(2) = "black"
	strA(3) = "red"

	'[3] �迭�� ���
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
		<font color="<%=strA(i)%>"> ���� <%=strA(i)%> </font><br>
	</td>
</tr>
<% 	Next %>
</table>