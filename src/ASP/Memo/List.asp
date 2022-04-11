<% 
Option Explicit
Dim objCon, strSql, objRs
strSql = "Select * From Memos Order By Num Desc"

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=MemoPwd;Persist Security Info=True;User ID=MemoUid;Initial Catalog=MemoDB;Data Source=.\SQLExpress")
    
Set objRs = objCon.Execute(strSql)
If objRs.BOF Or objRs.EOF Then
    Response.Write("입력된 데이터 없습니다.")
Else
    Do Until objRs.EOF
        Response.Write(objRs("Num") & _
            " " & objRs("Name") & "<br />")
        objRs.MoveNext()                    
    Loop
End If
Set objRs = Nothing        

objCon.Close()
Set objCon = Nothing
%>