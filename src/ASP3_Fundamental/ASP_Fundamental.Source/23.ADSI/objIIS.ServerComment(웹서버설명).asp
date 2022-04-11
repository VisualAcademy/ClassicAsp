<% Response.Expires = -1 %>
<%
If Request.ServerVariables("LOGON_USER") = "" Then
   Response.Status = "401 Authorization Required"
   Response.End
End If

Set objIIS = GetObject("IIS://" & Request.ServerVariables("SERVER_NAME") & "/W3SVC")
strComputerName = objIIS.ServerComment
If strComputerName = "" Then
	strComputerName = "(지정되지 않음)"
End If
%>
컴퓨터 이름: <b><%= Request.ServerVariables("SERVER_NAME") %></b><br>
웹 서버 설명: <b><%= strComputerName %></b>
