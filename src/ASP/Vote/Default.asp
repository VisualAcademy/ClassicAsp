<%
Option Explicit
Dim objCon, strSql, objRs
strSql = "Select Top 1 * From Vote Order By UID Desc"

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASPSample;Initial Catalog=ASPSample;Data Source=.\SQLEXPRESS")
Set objRs = objCon.Execute(strSql)

If objRs.BOF Or objRs.EOF Then
    Response.Write("진행중인 투표가 없습니다.")
Else
    Do Until objRs.EOF
        Response.Write("* " & objRs("Title") & "<br />")
        Call ShowItem(objRs("UID"), objRs("Allow_Multi")) ' 옵션 출력
        objRs.MoveNext()
    Loop
End If

objRs.Close()
Set objRs = Nothing
objCon.Close()
Set objCon = Nothing
%>
<%
Sub ShowItem(UID, Allow_Multi)
    Dim con, sql, rs
    sql = "Select * From Vote_Item Where Vote_ID = " & UID
    Set con = Server.CreateObject("ADODB.Connection")
    con.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASPSample;Initial Catalog=ASPSample;Data Source=(local)\SQLEXPRESS")
    Set rs = con.Execute(sql)
    If rs.BOF Or rs.EOF Then
        Response.Write("진행중인 투표가 없습니다.")
    Else
%>
    <form name="MyForm" action="Do_Vote.asp" method="get">
        <input type="hidden" name="Vote_ID" 
            value="<%= UID %>" />
<%
        Do Until rs.EOF
%>
        <input type="<% 
            If Allow_Multi = 0 Then
                Response.Write("radio")
            Else
                Response.Write("checkbox")
            End If
            %>" name="Vote_Item" 
        value="<%= rs("UID") %>" /><%= rs("Title") %><br />
<%
            rs.MoveNext()
        Loop
%>
        <input type="submit" value="투표하기" />
        <a href="Vote_Result.asp">설문결과</a>
    </form>
<%  
    End If
    rs.Close()
    Set rs = Nothing
    con.Close()
    Set con = Nothing
End Sub
%>