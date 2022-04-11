<%
	Response.Write( Request.ServerVariables("REMOTE_HOST") & "<br>")
	Response.Write( Request.ServerVariables("REMOTE_ADDR") & "<br>")
%>