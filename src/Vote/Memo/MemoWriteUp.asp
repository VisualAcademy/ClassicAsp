<%
Dim objCon
Dim strSql

Dim Name: Name = "������"
Dim Email: Email = "b@b.com"
Dim Title: Title = "�ȳ��ϼ���."
Dim PostIP: PostIP = Request.ServerVariables("REMOTE_HOST")

' �������ڿ� -> " & �ش纯�� & "
'strSql = "Insert Memos Values('ȫ�浿', 'h@h.com', '�ȳ�', GetDate(), '127.0.0.1')"
strSql = _
     "Insert Memos Values('" & Name & _
     "', '" & Email & _
     "', '" & Title & _
      "', GetDate(), '" & _
       PostIP & "')"

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASP;Initial Catalog=ASP;Data Source=.\SQLEXPRESS")
objCon.Execute(strSql)
objCon.Close()
Set objCon = Nothing
%>