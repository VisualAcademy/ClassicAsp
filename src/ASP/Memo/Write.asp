<%
Dim strSql
strSql = "Insert Memos(Name, Email, Title, PostIP) " & _
    " Values('ȫ�浿', 'h@h.com', '�ȳ�', '127.0.01') "
Dim objCon

'[1] �����ͺ��̽� ������ ���� Ŀ�ؼ� ��ü ����
Set objCon = Server.CreateObject("ADODB.Connection")
'[2] �����ͺ��̽� ����(����) : �����ͺ��̽� ���� ���ڿ� �ʿ�
objCon.Open("Provider=SQLOLEDB.1;Password=MemoPwd;Persist Security Info=True;User ID=MemoUid;Initial Catalog=MemoDB;Data Source=.\SQLExpress")
'[3] ��ɾ� ����
objCon.Execute(strSql)
'[4] �����ͺ��̽� �ݱ�
objCon.Close()
'[5] �����ͺ��̽� ��ü ���� ����
Set objCon = Nothing
%>
�����Ͱ� ����Ǿ����ϴ�.