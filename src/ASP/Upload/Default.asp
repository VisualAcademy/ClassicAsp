<%
Dim strSql: strSql = "Select * From Upload Order By Num Desc"

Dim objCon: Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("ConnectionString"))

Dim objRs: Set objRs = Server.CreateObject("ADODB.Recordset")
objRs.Open strSql, objCon

If objRs.BOF Or objRs.EOF Then
    Response.Write("�Էµ� �����Ͱ� �����ϴ�.")
Else
    Do Until objRs.EOF
        Response.Write(objRs("FileName") & "<br />")
        objRs.MoveNext()
    Loop 
End If

objRs.Close()
Set objRs = Nothing

objCon.Close()
Set objCon = Nothing
%>