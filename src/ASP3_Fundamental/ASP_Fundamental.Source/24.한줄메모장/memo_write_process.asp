<%
'--------------------------------------------------------------------
' Title : MyMemo Education Version
' Program Name : memo_write_process.asp
' Program Description : �޸�� �����ϴ� ���μ���
' Include Files : None
' Last Update : 2001.08.26. 02:20
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<%
'[1] ������ ����� ����
Option Explicit

'[2] ���� ����
Dim Name	'memo_write_form���� ���۵� �̸����� ������ ����
Dim Email	'memo_write_form���� ���۵� �̸��ϰ��� ������ ����
Dim Title		'memo_write_form���� ���۵� ����(�޸�)���� ������ ����
Dim PostDate	'MY�޸��� DB�� ��ϵ� ��¥�� ������ ����
Dim objDB		'Connection��ü�� �ν��Ͻ� ��ü�� ������ ����
Dim strSQL		'MY�޸��� DB�� �����͸� �����ϴ� SQL���� �����ϴ� ���� 

'[3] ���� �� ���� : Request��ü�� Form�÷������� �Ѿ�� ������(name, value)�� �޴� ����.
Name = Request.Form("name")
Email = Request.Form("email")
Title = Request.Form("title")
PostDate = Date()   'Now()

'[4] ���� �� ó�� : HTML�±� �� SQL Ư������ ���� ����(DB ���� ���� ����)
Title = Replace(Title,"'","''")   '����(Ȭ)����ǥ ó��
Title = Replace(Title,"&","&amp;")   '&(���ۻ���) ��ȣ ó��
Title = Replace(Title,"<","&lt;")   '< ��ȣ ó��
Title = Replace(Title,">","&gt;")   '> ��ȣ ó��

'[5] �����Ͱ� ����� �Ѿ�Դ��� �����(�׻� ����� �ϴ� ������ �鿩���� ��)
'Response.Write(Name & "," & Email & "," & Title & "," & PostDate)

'[6] �����ͺ��̽�(Connection��ü)�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")

'[7] �����ͺ��̽� ����(ODBC������ DSN������ SQL Server�� �α��� ����(uid/pwd) ���)
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'[8] �����ͺ��̽� ó��(�Է�) ���� �ۼ�(���� ����ǥ ��� ����)
'ASP ��¥ �Լ� ��� �Է�, "& _" �����ڷ� ����
'strSQL = "Insert my_memo(Name, Email, Title, PostDate) Values " & _
'"('" & Name & "','" & Email & "','" & Title & "'," & PostDate & ")"
'SQL ��¥ �Լ� ��� �Է�, "&" ���� �����ڷ� ����
strSQL = "Insert my_memo(Name, Email, Title, PostDate) Values "
strSQL = strSQL & "('" & Name & "','" & Email & "','" & Title & "',GetDate())"
'----------------------------------------------------------------------------------------
'Insert my_memo Values ('�ڿ���','abc@a.com','�޸�''����?','2001-11-13')
'Insert my_memo Values ('   �ڿ���  ','   abc@a.com   ','   �޸�''����?  ',' 2001-11-13 ')
'"Insert my_memo Values ('"  &  �ڿ��� & "','"  & abc@a.com &  "','"  & �޸�''����? & "','" 2001-11-13 & "')"
'strSQL = "Insert my_memo Values ('"  &  �ڿ��� & "','"  & abc@a.com &  "','"  & �޸�''����? & "','" 2001-11-13 & "')"
'strSQL = "Insert my_memo Values ('"  & Name & "','"  & Email &  "','"  & Title & "','" PostDate & "')"
'----------------------------------------------------------------------------------------

'[9] ���� ���� ����� : �Ʒ� ��°���� SQL Server �����м��⿡�� ������Ѻ���.
'Response.Write(strSQL)

'[10] �����ͺ��̽� ó��(�Է�) ����
objDB.Execute(strSQL)

'[11] �����ͺ��̽� �ݱ�
objDB.Close

'[12] �����ͺ��̽� ��ü�� �ν��Ͻ� ����
Set objDB = Nothing

'[13] �޸𸮽�Ʈ �������� �����̷�Ʈ - �׽�Ʈ�� ���� �Ʒ� �ڵ�� �ּ�ó��.
Response.Redirect "./memo_list.asp"
%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<CENTER>
			������ ���̽��� �ԷµǾ����ϴ�.
		</CENTER>
	</BODY>
</HTML>
