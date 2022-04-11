<%
Dim objCon, strSql

strSql = "Delete Memos Where Num = " & Request("Num")

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASP;Initial Catalog=ASP;Data Source=.\SQLEXPRESS")

objCon.Execute(strSql)

objCon.Close()
Set objCon = Nothing
%>
<div align="center">
<%=Request("Num") %>번 데이터가 삭제되었습니다.</div>
