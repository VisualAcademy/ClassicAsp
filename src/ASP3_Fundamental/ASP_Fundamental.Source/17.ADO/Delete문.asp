<%
'[1] ���� ����
Option Explicit
Dim objCon
Dim strSql
'[2] �����ͺ��̽��� �ν��Ͻ� ���� : Connection ������Ʈ
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] �����ͺ��̽� ����
objCon.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")

'[4] SQL�� ����
strSql = "Delete Memo Where Num = 5"
objCon.Execute(strSql)

'[5] �����ͺ��̽� �ݱ�
objCon.Close()
'[6] �����ͺ��̽��� �ν��Ͻ� ����
Set objCon = Nothing 
%>