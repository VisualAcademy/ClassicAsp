<%
Dim objCon, strSql, ID, User_Name, Pwd, Comment, User_IP

ID = Request("ID")
User_Name = Request("User_Name")
Pwd = Request("Pwd")
Comment = Request("Comment")

' 문자열 -> " & 변수 & "
strSql = "Update gBook Set User_Name = '" & User_Name & _
     "', Comment = '" & Comment & "' Where UID = " & ID
     
'Response.Write(strSql)
'Response.End()     

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=GuestBook;Initial Catalog=GuestBook;Data Source=172.16.25.99")
objCon.Execute(strSql)
objCon.Close()
Set objCon = Nothing

Response.Redirect("Default.asp")
%>