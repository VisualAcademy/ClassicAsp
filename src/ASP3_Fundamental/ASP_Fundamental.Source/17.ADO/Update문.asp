<%
'[1] ���� ����
Option Explicit
Dim objCon
'[2] �����ͺ��̽��� �ν��Ͻ� ���� : Connection ������Ʈ
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] �����ͺ��̽� ����
objCon.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")

'[4] SQL�� ����
objCon.Execute("Update Memo Set Email = 'h@h.com' Where Num = 2")

'[5] �����ͺ��̽� �ݱ�
objCon.Close()
'[6] �����ͺ��̽��� �ν��Ͻ� ����
Set objCon = Nothing 
%>