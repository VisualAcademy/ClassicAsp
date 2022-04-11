<%
Option Explicit
Dim Name: Name = Request("Name")
Dim Email: Email = Request("Email")
Dim Title: Title = Request("Title")
Dim IP: IP = Request.ServerVariables("REMOTE_ADDR")
Dim objCon: Dim strSql
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")
'순수 데이터 자리 : 홍길동 -> " & 변수 & "
strSql = "Insert Memo Values('" & Name & "', '" & Email & "', '" & Title & "', GetDate(), '" & IP & "')"
objCon.Execute(strSql)
objCon.Close()
Set objCon = Nothing
Response.Redirect("./MemoList.asp")
%>