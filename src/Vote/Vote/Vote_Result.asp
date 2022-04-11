<%
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
    sql = "Select Sum(Selected) From Vote_Item " & _ 
        "Where Vote_ID = " & UID & _
     ";Select * From Vote_Item Where Vote_ID = " & UID
    Set con = Server.CreateObject("ADODB.Connection")
    con.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASPSample;Initial Catalog=ASPSample;Data Source=(local)\SQLEXPRESS")
    Set rs = con.Execute(sql)
    If rs.BOF Or rs.EOF Then
        Response.Write("진행중인 투표가 없습니다.")
    Else
        Dim sum
        sum = rs(0)
        Set rs = rs.NextRecordset() ' 다음 SQL문 실행
        Do Until rs.EOF 
            Response.Write(rs("Title") & ", " & _
                "<img src='images/graph.gif' height='10' " & _ 
                "width='" & (rs("Selected")*100/sum)*2 & _
                "' />(" & rs("Selected") & ")<br />")
%>
    <%= rs("Title") %>, 
    <img src="images/graph.gif" height="10"
        width='<%= (rs("Selected")*100/sum)*2 %>' />
    (<%= rs("Selected") %>)<br />
<%
            rs.MoveNext()        
        Loop
    End If
    rs.Close()
    Set rs = Nothing
    con.Close()
    Set con = Nothing
End Sub
%>