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
'[!] ��������ǥ ó�� : Name, Title, Content ������...
Name = Replace(Name, "'", "''")
Title = Replace(Title, "'", "''")
Content = Replace(Content, "'", "'')

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("ConnectionString"))
	strSql = "Select Password From Upload Where Password = '" & Password & "' And Num = " & Num
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	objRs.Open strSql, objCon
		If objRs.BOF Or objRs.EOF Then
			'�н����� Ʋ���ٴ� ���� �Բ� �ڷ� ������
%>
			<script language="javascript">
			window.alert("��й�ȣ�� Ʋ���ϴ�.");
			window.history.go(-1);
			</script>
<%
		Else
			'������Ʈ �� ���� �� View �������� ��ȣ������ �̵�
			'����ó�� : ���������� -> " & ���� & "
			strSql = "Update Upload Set Name = '" & Name & "', Email = '" & Email & "', Title = '" & Title & "', Content = '" & Content & "', Encoding = '" & Encoding & "', Homepage = '" & Homepage & "', ModifyDate = GetDate(), ModifyIP = '" & Request.ServerVariables("REMOTE_HOST") & "' Where Num = " & Num & " And Password = '" & Password & "'"
			objCon.Execute(strSql)
			Response.Redirect("BoardView.asp?Num=" & Num)'//�̵�
		End If
	objRs.Close()
	Set objRs = Nothing
objCon.Close()
Set objCon = Nothing
%>










