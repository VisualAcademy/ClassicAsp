<%
Dim objCon: Dim objRs: Dim strSql
Dim Num: Num = Request("Num")
Dim Name: Name = Request("Name")
Dim Email: Email = Request("Email")
Dim Homepage: Homepage = Request("Homepage")
Dim Title: Title = Request("Title")
Dim Content: Content = Request("Content")
Dim Encoding: Encoding = Request("Encoding")
Dim Password: Password = Request("Password")
'[!] ����(Ȭ)����ǥ ó�� : '(��������ǥ �ϳ�->''(��������ǥ �ΰ�)
Name = Replace(Name, "'", "''")
Title = Replace(Title, "'", "''")
Content = Replace(Content, "'", "''")
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
strSql = "Select * From Basic Where Num = " & Num & " And Password = '" & Password & "'"
Set objRs = objCon.Execute(strSql)
If objRs.BOF Or objRs.EOF Then
	'//�� �κп��� �ڹٽ�ũ��Ʈ�� �ڷ� ������//
	Response.End()
Else
	'// ������ ������ -> " & ���� & "
	strSql = "Update Basic Set Name = '" & Name & "', Email = '" & Email & "', Title = '" & Title & "', Content = '" & Content & "', Encoding = '" & Encoding & "', Homepage = '" & Homepage & "', ModifyDate = GetDate(), ModifyIP = '" & Request.ServerVariables("REMOTE_HOST") & "' Where Num = " & Num & ""
	objCon.Execute(strSql)	
End If
objCon.Close()
Set objCon = Nothing
'//���� �� ��� �̵�
Response.Redirect("BoardView.asp?Num=" & Request("Num"))
%>