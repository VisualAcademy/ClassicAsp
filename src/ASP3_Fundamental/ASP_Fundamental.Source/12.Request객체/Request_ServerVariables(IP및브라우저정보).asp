<%
	'Request��ü�� ServerVariables()�÷����� ���������� ����Ѵ�.
	'IP�ּ� ���
	Response.Write(Request.ServerVariables("REMOTE_HOST") & "<br>")
	'IP�ּ� ���
	Response.Write(Request.ServerVariables("REMOTE_ADDR") & "<br>")
%>
������ ���� : <%=Request.ServerVariables("HTTP_USER_AGENT")%><br>
������Ʈ : <%=Request.ServerVariables("SERVER_PORT")%><br>
���.
