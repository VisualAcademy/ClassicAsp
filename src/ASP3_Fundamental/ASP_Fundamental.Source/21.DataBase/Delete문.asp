<%
	'[1] ������ ����� ����
	Option Explicit

	'[2] ���� ����
	Dim objDB, strSQL

	'[3] �����ͺ��̽��� �ν��Ͻ�(��ü)�� ����
	Set objDB = Server.CreateObject("ADODB.Connection")

	'[4] �����ͺ��̽� ����(Open() �޼ҵ� ��� : DSN, UID, PWD)
	objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

	'[5] SQL ���ڿ� �ۼ�
 	strSQL = "Delete my_memo"   'my_memo���̺��� ��� ������ ����

	'[6] SQL ���ڿ� ����(Execute()�޼ҵ� ���)
	objDB.Execute strSQL

	'[7] �����ͺ��̽� �ݱ�
	objDB.Close

	'[8] �����ͺ��̽��� �ν��Ͻ� ���� ����
	Set objDB = Nothing
%>
<center>��� �ڷᰡ �����Ǿ����ϴ�.</center>
