<%
'[1] ���� ����
Option Explicit
Dim objCon
'[2] �����ͺ��̽��� �ν��Ͻ� ���� : Connection ������Ʈ
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] �����ͺ��̽� ����
objCon.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")
'[4] SQL�� ����
objCon.Execute("Insert Memo Values('ȫ�浿', 'hong@h.com', 'ȫ�浿�Դϴ�.', GetDate(), '127.0.0.1')")
'[5] �����ͺ��̽� �ݱ�
objCon.Close()
'[6] �����ͺ��̽��� �ν��Ͻ� ����
Set objCon = Nothing 
%>