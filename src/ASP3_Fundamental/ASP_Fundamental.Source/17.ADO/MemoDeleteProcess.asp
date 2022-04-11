<%
Dim Num: Num = Request("Num")
Dim Password: Password = Request("Password")

If Password = "1234" Then
	Call DeleteArticle()
	Response.Redirect("./MemoList.asp")
Else
%>
	<SCRIPT LANGUAGE="JavaScript">
	window.alert("해킹하려구???");
	history.go(-1);
	</SCRIPT>
<%
End If
%>
<%
Sub DeleteArticle()
	Set objCon = Server.CreateObject("ADODB.Connection")
	objCon.Open("Provider=SQLOLEDB.1;Password=memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")
	'// 1 -> " & 변수 & "
	objCon.Execute("Delete Memo Where Num = " & Num)
	objCon.Close()
	Set objCon = Nothing
End Sub
%>