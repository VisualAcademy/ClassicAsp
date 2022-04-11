<%
Dim strSql: strSql = "Select * From Memos"

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASP;Initial Catalog=ASP;Data Source=.\SQLEXPRESS")

'Set objRs = objCon.Execute(strSql)
Set objRs = Server.CreateObject("ADODB.Recordset")
objRs.Open strSql, objCon
If objRs.BOF Or objRs.EOF Then

Else
    Call ShowRecordSet(objRs)
End If
objRs.Close()
Set objRs = Nothing

objCon.Close()
Set objCon = Nothing
%>

<%
Sub ShowRecordSet(ByRef objRs)
%>
    <table border="1" width="100%">
        <tr>
            <td>번호</td><td>제목</td>
        </tr>
    <%
        Do Until objRs.EOF
    %>
        <tr>
            <td><%= objRs("Num") %></td>
            <td><%= objRs.Fields("Name") %></td>
        </tr>                
    <%
            objRs.MoveNext()
        Loop
    %>
    </table>
<%
End Sub
%>