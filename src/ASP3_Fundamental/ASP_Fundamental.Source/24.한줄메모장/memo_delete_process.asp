<%
'--------------------------------------------------------------------
' MyMemo Education Version
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<%
Option Explicit

'���� ����
Dim objDB, strSQL, Num, Passwd

Num = Request.Form("Num")
Passwd = Request.Form("passwd")

If Num <> "" AND Passwd = "1234" Then
	'�����ͺ��̽�(connection��ü)�� �ν��Ͻ� ����
	Set objDB = Server.CreateObject("ADODB.Connection")

	'�����ͺ��̽� ����
	objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

	'�����ͺ��̽� ó��(����) ���� �ۼ�
	strSQL = "Delete my_memo Where Num =" & Num
        
	'�����ͺ��̽� ó��(�Է�) ����
	objDB.Execute(strSQL)

	'�޸𸮽�Ʈ �������� �����̷�Ʈ - �׽�Ʈ�� ���� �Ʒ� �ڵ�� �ּ�ó��.
	Response.Redirect "./memo_list.asp"

Else
%>
<SCRIPT language="javascript">
	history.go(-1);
</SCRIPT>
<%
End If
%>
<%
	'�����ͺ��̽� �ݱ�
	objDB.Close

	'�����ͺ��̽� ��ü�� �ν��Ͻ� ����
	Set objDB = Nothing
%>
