<%
'--------------------------------------------------
' Title : Myado Education Version
' Program Name : ado_modify_process.asp
' Program Description : ���������� ���۵� �� Update������ ����
' Include Files : ado_function.asp
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!--METADATA TYPE= "typelib" NAME= "ADODB Type Library" FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll" -->

<!-- HTML�±� �� SQL Ư������  ���� ���� ���� ����� �����Լ� ȣ�� -->
<!-- #include file="./ado_function.asp" -->

<%
'Option Explicit   '�� ������ �ּ�ó���� ������ ��Ŭ��� ���Ͽ��� ���������� ���߱⶧��(��������)

'���� ����
Dim Num, Name, Email, Title, ModifyDate, ModifyIP, Encoding, Passwd, Content, objDB, strSQL
Dim strDB


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

'--------------------------------------------------------------------
'[1] ���ڵ� �� ��ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet")

'[2] �����ͺ��̽� ���� ���ڿ� ����
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=XPPLUSDOTNET"

'[3] �Ѱܿ� �۹�ȣ�� �ش�Ǵ� �н����� ��������
strSQL = "Select * From my_ado Where Num = " & Num
'Response.Write strSQL

'[4] ���ڵ� �� ����
objRs.Open strSQL, strDB, adOpenStatic, adLockPessimistic, adCmdText

If objRS.Fields("Passwd") = Passwd Then

With objRS
	
	.Fields("Name") = Name
	.Fields("Email") = Email
	.Fields("Title") = Title
	.Fields("ModifyDate") = Now()
	.Fields("Encoding") = Encoding
	.Fields("ModifyIP") = ModifyIP
	.Fields("Content") = Content

	.Update

	.Close

End With


	'���ڵ�� ��ü�� �ν��Ͻ� ����
	Set objRS = Nothing
'--------------------------------------------------------------------

	'�Խ��� ����Ʈ �������� �����̷�Ʈ - �׽�Ʈ�� ���� �Ʒ� �ڵ�� �ּ�ó��.
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
