<%
'--------------------------------------------------------------------
' Title : ���ڵ�� ��ü�� �̿��� �ڷ�����.
' Program Name : ado_write_process.asp
' Program Description : ADO ��� ����
' Include Files : ADO constants include file for VBScript
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>

<!-- HTML�±� �� SQL Ư������  ���� ���� ���� ����� �����Լ� ȣ�� -->
<!-- #include file="./ado_function.asp" -->

<!-- ------------------------------------------- -->
<!-- #include file="./scripts/adovbs.inc" -->
<!-- ------------------------------------------- -->

<%
'Option Explicit   '�� ������ �ּ�ó���� ������ ��Ŭ��� ���Ͽ��� ���������� ���߱⶧��(��������)

'���� ����
Dim Name, Email, Title, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Content, objDB, objRS, strSQL

'Request��ü�� Form�÷������� �Ѿ�� ������(name, value)�� �޴� ����.
Name = Request("Name")
Email = Request("Email")
Title = Request("Title")
PostDate = Date()
ModifyDate = Date() 'ó�� �۾����� �ش� ��¥ �Է�.
ReadCount = 0
PostIP = CStr(Request.ServerVariables("REMOTE_HOST"))
ModifyIP = CStr(Request.ServerVariables("REMOTE_HOST")) 'ó�� �۾����� �ش� IP �Է�.
Encoding = Request.Form("Encoding")
Passwd = Request.Form("Password")
Content = Request.Form("Content")

'HTML�±� �� SQL Ư������  ���� ����.-----------------------------------
	Title = ConvertToHTML(Title)   '���� HTML������
	Email = ConvertToHTML(Email)   'Email�� HTML������
	
	Title = ConvertToSQL(Title)   'Title�� ��������ǥ ó��
	Email = ConvertToSQL(Email)   'Email�� ��������ǥ ó��
	Content = ConvertToSQL(Content)   'Content�� ��������ǥ ó��

	'���� �� ���ڵ��� HTML���� Text������ ���� ���� ���� ����.
	Select Case Encoding

		Case "HTML"   'HTML�� ó���ؼ� �����ٶ�
	
			Content = CStr(Content)
		
		Case "Text"   'HTML �ҽ��� �״�� ������ ��

			Content = CStr(Content)
			Content = ConvertToHTML(Content)   

	End Select
'HTML�±� �� SQL Ư������  ���� ����.-----------------------------------

'1. �����Ͱ� ����� �Ѿ�Դ��� �����
'Response.Write(Name & "," & Email & "," & Title & "," & PostDate & "," & ModifyDate & "," & ReadCount & "," & PostIP & "," & ModifyIP & "," & Encoding & "," & Passwd& "," & Content)



'--------------------------------------------------------------------
' ���ڵ� �� ��ü�� �̿��� ������ ����
'--------------------------------------------------------------------
'[1] ���ڵ� �� ��ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet")

'[2]�����ͺ��̽� ���� ���ڿ� ����
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=ZZAZIPKIDOTCOM"

'[3] ���ڵ� �� ����(���⼭�� ù��° ���ڰ��� SQL�� ��ſ� �ش� ���̺� �̸� ����)
objRS.Open "my_ado", strDB, adOpenStatic, adLockOptimistic, adCmdTable

	objRS.AddNew   '���ο� ���ڵ带 ������ ���� ����(Connection��ü�� Insert���� ����)

	objRS.Fields("Name") = Name
	objRS.Fields("Email") = Email
	objRS.Fields("Title") = Title
	objRS.Fields("PostDate") = Now()
	objRS.Fields("ModifyDate") = Now()
	objRS.Fields("ReadCount") = 0
	objRS.Fields("PostIP") = CStr(Request.ServerVariables("REMOTE_HOST"))
	objRS.Fields("ModifyIP") = CStr(Request.ServerVariables("REMOTE_HOST"))
	objRS.Fields("Encoding") = Encoding
	objRS.Fields("Passwd") = Passwd
	objRS.Fields("Content") = Content

	objRS.Update   '������ ������ ���� ����

	'���ڵ� �� ��ü�� �ν��Ͻ� ����
	Set objRS = Nothing
'--------------------------------------------------------------------

'3. �Խ��Ǹ���Ʈ �������� �����̷�Ʈ - �׽�Ʈ�� ���� �Ʒ� �ڵ�� �ּ�ó��.
Response.Redirect "./ado_list.asp"

%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<br>������ ���̽��� �ԷµǾ����ϴ�.
	</BODY>
</HTML>