<%
	'[1] ������ ����� ����
	Option Explicit

	'[2] ���� ����
	Dim objDB, strSQL
	Dim Name: Name="ȫ�浿"
	Dim Email: Email="h@h.com"
	Dim Title: Title="ȫ�浿�Դϴ�."
	Dim PostIP: PostIP="127.0.0.1"

	'[3] �����ͺ��̽��� �ν��Ͻ�(��ü)�� ����
	Set objDB = Server.CreateObject("ADODB.Connection")

	'[4] �����ͺ��̽� ����(Open() �޼ҵ� ��� : DSN, UID, PWD)
	objDB.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")

	'[5] SQL ���ڿ� �ۼ�
	'ȫ�浿 -> " & ���� & "
	strSQL = "Insert Memo(Name, Email, Title, PostDate, PostIP) Values('" & Name & "','" & Email & "','" & Title & "',GetDate(),'" & PostIP & "')"

	'[6] SQL�� �����
	Response.Write strSQL

	'[7] SQL ���ڿ� ����(Execute()�޼ҵ� ���)
	objDB.Execute strSQL

	'[8] �����ͺ��̽� �ݱ�
	objDB.Close

	'[9] �����ͺ��̽��� �ν��Ͻ� ���� ����
	Set objDB = Nothing
%>
<center>�ԷµǾ����ϴ�.</center>
