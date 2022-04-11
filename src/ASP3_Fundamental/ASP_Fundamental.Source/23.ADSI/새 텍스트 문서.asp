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
<p>
<form action="<%= Request.ServerVariables("PATH_INFO") %>" method="post">

<b><font color="#ff0000">���� ������ ������Ʈ ���</font></b>

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