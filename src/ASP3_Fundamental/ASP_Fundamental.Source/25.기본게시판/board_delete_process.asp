<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_delete_process.asp
' Program Description : ���۵� �н����尡 ������ �ش�� �����
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
'������ ����� ����
Option Explicit
Dim Num, Passwd, objDB, objRS, strSQL

Num = Request("Num")
Passwd = Request("Passwd")

'�����ͺ��̽�(connection��ü)�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")

'�����ͺ��̽� ����
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'���ڵ�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet")

'�Ѱܿ� �۹�ȣ�� �ش�Ǵ� �н����� ��������
strSQL = "Select Passwd From my_board Where Num = " & Num
'Response.Write strSQL

'���ڵ� �� ����
objRS.Open strSQL, objDB, 1


If objRS.Fields("Passwd") = Passwd Then
	'�����ͺ��̽� ó��(����) ���� �ۼ�
	strSQL = "Delete my_board Where Num =" & Num
        
	'�����ͺ��̽� ó��(�Է�) ����
	objDB.Execute(strSQL)

	'���ϸ���Ʈ �������� �����̷�Ʈ - �׽�Ʈ�� ���� �Ʒ� �ڵ�� �ּ�ó��.
	Response.Redirect "./board_list.asp"

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

<%
	'���ڵ�� �ݱ�
	objRS.Close

	'�����ͺ��̽� �ݱ�
	objDB.Close

	'���ڵ�� ��ü�� �ν��Ͻ� ����
	Set objRS = Nothing

	'�����ͺ��̽� ��ü�� �ν��Ͻ� ����
	Set objDB = Nothing
%>