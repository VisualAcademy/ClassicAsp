<%
	'[1] ������ ����� ����
	Option Explicit

	'[2] ���� ����
	Dim objDB, strSQL

	'[3] �����ͺ��̽��� �ν��Ͻ�(��ü)�� ����
	Set objDB = Server.CreateObject("ADODB.Connection")
	
	'[4] �����ͺ��̽� ����(Open() �޼ҵ� ��� : DSN, UID, PWD)
	objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")
	
	'[5] SQL ���� ���ڿ� �ۼ�
	strSQL = "Update my_memo Set Email = 'bbb@bbb.com'"

	'[6] SQL ���ڿ� ����(Execute()�޼ҵ� ���)
	objDB.Execute strSQL

	'[7] �����ͺ��̽� �ݱ�
	objDB.Close

	'[8] �����ͺ��̽��� �ν��Ͻ� ���� ����
	Set objDB = Nothing
%>
<center>�����Ǿ����ϴ�.</center>
