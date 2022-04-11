<%
Option Explicit
Dim objCon, strSql
strSql = "Insert Memos Values('홍길동','h@h.com'," & _ 
    " '안녕', GetDate(), '127.0.0.')"

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASP;Initial Catalog=ASP;Data Source=.\SQLEXPRESS")

objCon.Execute(strSql)

objCon.Close()
Set objCon = Nothing
%>
<div align="center">저장되었습니다.</div>
