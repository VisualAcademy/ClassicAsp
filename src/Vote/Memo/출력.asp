<%
Option Explicit
Dim objCon, strSql, objRs
strSql = "Select * From Memos Order By Num Desc"
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASP;Initial Catalog=ASP;Data Source=.\SQLEXPRESS")
Set objRs = objCon.Execute(strSql)
If objRs.BOF Or objRs.EOF Then
    Response.Write("데이터가 없습니다.")
Else
    Call ShowRecordSet(objRs)
End If
Set objRs = Nothing
objCon.Close()
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(ByRef objRs)
%>
    <table border="1" width="500">
    <%
        Do Until objRs.EOF
            Response.Write("<tr><td>")
            Response.Write(objRs.Fields("Name") & "<br />")
            Response.Write("</td></tr>")
            objRs.MoveNext()
        Loop
    %>
    </table>
<%
End Sub
%>
