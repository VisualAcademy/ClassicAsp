<% Response.Expires = -1 %>
<%
If Request.ServerVariables("LOGON_USER") = "" Then
   Response.Status = "401 Authorization Required"
   Response.End
End If

Set objIIS = GetObject("IIS://" & Request.ServerVariables("SERVER_NAME") & "/W3SVC")
strComputerName = objIIS.ServerComment
If strComputerName = "" Then
	strComputerName = "(�������� ����)"
End If
%>
��ǻ�� �̸�: <b><%= Request.ServerVariables("SERVER_NAME") %></b><br>
�� ���� ����: <b><%= strComputerName %></b>
