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
	strSQL = "Insert my_memo Values('ȫ�浿','hong@hong.com','ȫ�浿�Դϴ�.',GetDate())"

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
