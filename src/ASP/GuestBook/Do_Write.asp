<%
Dim objCon, strSql, User_Name, Pwd, Comment, User_IP

User_Name = Request("User_Name")
Pwd = Request("Pwd")
Comment = Request("Comment")
User_IP = Request.ServerVariables("REMOTE_ADDR")

' 매개변수처리 : 순수문자열 -> " & 해당변수 & "
strSql = "Insert gBook Values('" & User_Name & "', '" & _ 
    Comment & "', GetDate(), '" & User_IP & _
    "', '" & Pwd & "')"
'Response.Write(strSql)
'Response.End()

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=GuestBook;Initial Catalog=GuestBook;Data Source=172.16.25.99")
objCon.Execute(strSql)
objCon.Close()
Set objCon = Nothing

Response.Redirect("Default.asp")
%>