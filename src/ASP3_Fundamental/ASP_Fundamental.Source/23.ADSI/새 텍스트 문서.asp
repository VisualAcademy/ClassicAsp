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
<p>
<form action="<%= Request.ServerVariables("PATH_INFO") %>" method="post">

<b><font color="#ff0000">현재 서버의 웹사이트 목록</font></b>

<select name="IIsWebServer">

<%
For Each IISObjects In objIIS

	If IISObjects.Class = "IIsWebServer" Then
%>
	<option value="<%= IISObjects.Name %>"><%= IISObjects.Name %>: <%= IISObjects.ServerComment %></option>

<%
	End IF

Next
%>
</select>
<p>
<%
Set objIIS = Nothing
%>