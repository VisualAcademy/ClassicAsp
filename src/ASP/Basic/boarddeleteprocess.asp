<%
'--------------------------------------------------
' Title : Basic ����
' Program Name : boarddeleteprocess.asp
' Program Description : ����� ��
' QueryString : delete.asp���� Num=1, Password=1������ ���� ���۵�...
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
Option Explicit
Dim Num: Num = Request("Num")
Dim Password: Password = Request("Password")
Dim objCon: Dim objRs: Dim strSql
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("CONNECTIONSTRING"))
Set objRs = Server.CreateObject("ADODB.RecordSet")
strSql = "Select Password From Basic" & _
		 " Where Num = "
strSql = strSql & Num
objRs.Open strSql, objCon
If objRs("Password") = Password Then '�н����尡 ���ٸ�,
	strSql = "Delete Basic Where Num = " & Num
	objCon.Execute(strSql)
	Response.Redirect("./boardlist.asp")
Else
%>
	<SCRIPT LANGUAGE="JavaScript">
	window.alert("��ȣ�� Ʋ���ϴ�.\n��ȣ�� Ȯ���ϼ���.");
	history.go(-1);
	</SCRIPT>
<%
End If
objRs.Close(): objCon.Close()
Set objRs = Nothing: Set objCon = Nothing
%>