<%
Dim objCon, strSql, objRs
strSql = "Select * From gBook Where UID = " & Request("ID")
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=GuestBook;Initial Catalog=GuestBook;Data Source=172.16.25.99")
Set objRs = Server.CreateObject("ADODB.Recordset")
objRs.Open strSql, objCon
If objRs.BOF Or objRs.EOF Then
    Response.Write("�����Ͱ� �����ϴ�.")
Else
%>
    <h1>���� �� ����</h1>
    <form name="MyForm" action="Do_Update.asp" method="post">
    <input type="hidden" name="ID" value="<%= Request("ID") %>" />
    �̸� : <input type="text" name="User_Name" value="<%=objRs("User_Name") %>" /><br />
    ��ȣ : <input type="password" name="Pwd" /><br />
    <textarea name="Comment" rows="10" cols="40"><%= objRs("Comment") %></textarea><br />
    <input type="submit" value="����" />
    </form>
<%
End If
objRs.Close()
Set objRs = Nothing
objCon.Close()
Set objCon = Nothing
%>
