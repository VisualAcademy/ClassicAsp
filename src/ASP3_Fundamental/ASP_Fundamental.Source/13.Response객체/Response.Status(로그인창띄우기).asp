<%
If Request.ServerVariables("LOGON_USER") = "" Then
   Response.Status = "401 Authorization Required"
   Response.End
End If

Response.Write("로그인 되었습니다.")
%>
