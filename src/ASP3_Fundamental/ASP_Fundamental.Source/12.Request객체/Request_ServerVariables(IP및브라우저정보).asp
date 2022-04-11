<%
	'Request객체의 ServerVariables()컬렉션은 서버변수를 출력한다.
	'IP주소 출력
	Response.Write(Request.ServerVariables("REMOTE_HOST") & "<br>")
	'IP주소 출력
	Response.Write(Request.ServerVariables("REMOTE_ADDR") & "<br>")
%>
브라우저 정보 : <%=Request.ServerVariables("HTTP_USER_AGENT")%><br>
서버포트 : <%=Request.ServerVariables("SERVER_PORT")%><br>
등등.
