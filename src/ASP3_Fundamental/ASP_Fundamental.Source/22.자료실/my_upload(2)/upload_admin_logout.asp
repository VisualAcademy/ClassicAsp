<%
'--------------------------------------------------
' Title : Myupload Education Version
' Program Name : upload_admin_logout.asp
' Program Description : ������ �α׾ƿ� �ϱ�
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
If Session("ID") = "Admin" Then
	Session("ID") = ""
End If

Response.Redirect "./upload_list.asp"
%>
