<%
Option Explicit
Dim Name: Name = Request("Name")
Dim Email: Email = Request("Email")
Dim Title: Title = Request("Title")
Dim IP: IP = Request.ServerVariables("REMOTE_ADDR")
Dim objCon: Dim strSql
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASP;Initial Catalog=ASP;Data Source=.\SQLEXPRESS")
'���� ������ �ڸ� : ȫ�浿 -> " & ���� & "
strSql = "Insert Memos Values('" & Name & "', '" & Email & "', '" & Title & "', GetDate(), '" & IP & "')"
objCon.Execute(strSql)
objCon.Close()
Set objCon = Nothing
Response.Redirect("./MemoList.asp")
%> 