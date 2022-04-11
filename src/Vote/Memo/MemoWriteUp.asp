<%
Dim objCon
Dim strSql

Dim Name: Name = "지리산"
Dim Email: Email = "b@b.com"
Dim Title: Title = "안녕하세요."
Dim PostIP: PostIP = Request.ServerVariables("REMOTE_HOST")

' 순수문자열 -> " & 해당변수 & "
'strSql = "Insert Memos Values('홍길동', 'h@h.com', '안녕', GetDate(), '127.0.0.1')"
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