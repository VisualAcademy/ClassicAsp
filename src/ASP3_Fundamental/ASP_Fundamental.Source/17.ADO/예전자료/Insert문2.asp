<%
'[1] ���� ����
Option Explicit
Dim objCon
'[2] �����ͺ��̽��� �ν��Ͻ� ����
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] �����ͺ��̽� ���� : �����ͺ��̽����Ṯ�ڿ� ����
objCon.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")
'[4] SQL�� ����
objCon.Execute("Insert Memo Values('ȫ�浿', 'hong@hong.com', '�׽�Ʈ�Դϴ�.', GetDate())")
'[5] �����ͺ��̽� �ݱ�
objCon.Close()
'[6] �����ͺ��̽��� �ν��Ͻ� ���� ����
Set objCon = Nothing
%>