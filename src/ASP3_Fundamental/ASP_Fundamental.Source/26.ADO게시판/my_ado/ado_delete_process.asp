<%
'--------------------------------------------------
' Title : Myado Education Version
' Program Name : ado_delete_process.asp
' Program Description : ���۵� �н����尡 ������ �ش�� �����
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!-- #include file="./scripts/adovbs.inc" -->

<%
'������ ����� ����
'Option Explicit
Dim Num
Dim Passwd
Dim objRS
Dim strDB
Dim strSQL

Num = Request("Num")
Passwd = Request("Passwd")

'[1] ���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'[2] �����ͺ��̽� ���� ���ڿ� ����
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=XPPLUSDOTNET"

'[3] �Ѱܿ� �۹�ȣ�� �ش�Ǵ� �н����� ��������
strSQL = "Select * From my_ado Where Num = " & Num
'Response.Write strSQL

'[4] ���ڵ�� ����
objRs.Open strSQL, strDB, adOpenStatic, adLockPessimistic, adCmdText
'objRS.Open strSQL, Application("Connection_string"), adOpen...........   <--- global.asa�� �̸� ���ø����̼� ������ �������� ���.

If objRS.Fields("Passwd") = Passwd Then
	'�����ͺ��̽� ó��(����) ���� �ۼ�

	objRS.Delete   '���� �ҷ��� ���ڵ�� ����

	objRS.Update   '���� ����.

	'���ڵ�� �ݱ�
	objRS.Close

	'���ڵ�� ��ü�� �ν��Ͻ� ����
	Set objRS = Nothing

	'���ϸ���Ʈ �������� �����̷�Ʈ - �׽�Ʈ�� ���� �Ʒ� �ڵ�� �ּ�ó��.
	Response.Redirect "./ado_list.asp"

Else
%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio.NET 7.0">
	</HEAD>
	<BODY>
		<script language="javascript">
			window.alert("��й�ȣ�� ��ġ���� �ʽ��ϴ�");
			history.go(-1);
		</script>
	</BODY>
</HTML>
<%
End If
%>

