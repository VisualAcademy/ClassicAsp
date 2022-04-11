<%
Option Explicit
Dim objCon, strSql, objRs
strSql = "Select * From Memos Where Num = " & Request("Num")
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASP;Initial Catalog=ASP;Data Source=.\SQLEXPRESS")

Set objRs = Server.CreateObject("ADODB.Recordset")
objRs.Open strSql, objCon
If objRs.BOF Or objRs.EOF Then
    Response.Write("데이터가 없습니다.")
Else
    Do Until objRs.EOF
        Response.Write(objRs("Name") & "<br />")
        objRs.MoveNext()
    Loop
End If
objRs.Close()
Set objRs = Nothing

objCon.Close()
Set objCon = Nothing
%>