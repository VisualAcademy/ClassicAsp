<%
'--------------------------------------------------
' Title : Myupload Education Version
' Program Name : upload_modify_process.asp
' Program Description : ���������� ���۵� �� Update������ ����
' Include Files : upload_function.asp
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!-- HTML�±� �� SQL Ư������  ���� ���� ���� ����� �����Լ� ȣ�� -->
<!-- #include file="./upload_function.asp" -->

<%
'Option Explicit   '�� ������ �ּ�ó���� ������ ��Ŭ��� ���Ͽ��� ���������� ���߱⶧��(��������)

'���� ����
Dim Num, Name, Email, Title, ModifyDate, ModifyIP, Encoding, Passwd, Content, objDB, strSQL

'Request��ü�� Form�÷������� �Ѿ�� ������(name, value)�� �޴� ����.
Num = Request.Form("Num")
Name = Request.Form("Name")
Email = Request.Form("Email")
Title = Request.Form("Title")
ModifyDate = Date() 'ó�� �۾����� �ش� ��¥ �Է�.
Passwd = Request.Form("password")
ModifyIP = Request.ServerVariables("REMOTE_HOST") 'ó�� �۾����� �ش� IP �Է�.
Encoding = Request.Form("Encoding")
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

'���ڵ�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet")

'�Ѱܿ� �۹�ȣ�� �ش�Ǵ� �н����� ��������
strSQL = "Select Passwd From my_upload Where Num = " & Num
'Response.Write strSQL

'���ڵ� �� ����
objRS.Open strSQL, objDB, 1

If objRS.Fields("Passwd") = Passwd Then
	'�����ͺ��̽� ó��(����) ���� �ۼ�
	strSQL2 = "Update my_upload Set Name = '" & Name & "', Email = '" & Email & "', Title = '" & Title & "', ModifyDate = '" & ModifyDate & "', Encoding = '" & Encoding & "', ModifyIP = '" & ModifyIP & "', Content = '" & Content & "' Where Num = " & Num
	'Response.Write strSQL2
        
	'�����ͺ��̽� ó��(�Է�) ����
	objDB.Execute(strSQL2)

	'���ڵ�� �ݱ�
	objRS.Close

	'�����ͺ��̽� �ݱ�
	objDB.Close

	'���ڵ�� ��ü�� �ν��Ͻ� ����
	Set objRS = Nothing

	'�����ͺ��̽� ��ü�� �ν��Ͻ� ����
	Set objDB = Nothing

	'���ϸ���Ʈ �������� �����̷�Ʈ - �׽�Ʈ�� ���� �Ʒ� �ڵ�� �ּ�ó��.
	Response.Redirect "./upload_list.asp"

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
