<%
If Request.ServerVariables("LOGON_USER") = "" Then
   Response.Status = "401 Authorization Required"
   Response.End
End If

Response.Write("�α��� �Ǿ����ϴ�.")
%>
