<%
Response.Write("���� ������ IP�ּ� : " & _
	Request.ServerVariables("REMOTE_HOST") & "<br>")
Response.Write("���� ������ IP�ּ� : " & _
	Request.ServerVariables("REMOTE_ADDR") & "<br>")
%>