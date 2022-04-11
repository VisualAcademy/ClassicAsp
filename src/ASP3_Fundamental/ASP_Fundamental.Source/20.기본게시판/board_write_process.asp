<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_write_process.asp
' Program Description : �Է������� ���۵� �� Insert������ ����
' Include Files : board_function.asp
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!-- HTML�±� �� SQL Ư������  ���� ���� ���� ����� �����Լ� ȣ�� -->
<!-- #include file="./board_function.asp" -->

<%
'Option Explicit   '�� ������ �ּ�ó���� ������ ��Ŭ��� ���Ͽ��� ���������� ���߱⶧��(��������)

'���� ����
Dim Name, Email, Title, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Content, objDB, strSQL

'Request��ü�� Form�÷������� �Ѿ�� ������(name, value)�� �޴� ����.
Name = Request.Form("Name")
Email = Request.Form("Email")
Title = Request.Form("Title")
PostDate = Date()
ModifyDate = Date() 'ó�� �۾����� �ش� ��¥ �Է�.
ReadCount = 0
PostIP = Request.ServerVariables("REMOTE_HOST")
ModifyIP = Request.ServerVariables("REMOTE_HOST") 'ó�� �۾����� �ش� IP �Է�.
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

'�����ͺ��̽�(connection��ü)�� �ν��Ͻ� ����
Set objDB = Server.Createobject("ADODB.Connection")

'�����ͺ��̽� ����
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'�����ͺ��̽� ó��(�Է�) ���� �ۼ�
strSQL = "Insert my_board(Name, Email, Title, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Content) Values " & _
         "('" & Name & "','" & Email & "','" & Title & "','" & PostDate & "','" & _
         ModifyDate & "'," & ReadCount & ",'" & PostIP & "','" & ModifyIP & "','" & _
		 Encoding & "','" & Passwd & "','" & Content & "')"
        
'2. ���� ���� �����
'Response.Write(strSQL)

'�����ͺ��̽� ó��(�Է�) ����
objDB.Execute(strSQL)

'�����ͺ��̽� �ݱ�
objDB.Close

'�����ͺ��̽� ��ü�� �ν��Ͻ� ����
Set objDB = Nothing

'3. �Խ��Ǹ���Ʈ �������� �����̷�Ʈ - �׽�Ʈ�� ���� �Ʒ� �ڵ�� �ּ�ó��.
Response.Redirect "./board_list.asp"

%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<br>������ ���̽��� �ԷµǾ����ϴ�.
	</BODY>
</HTML>