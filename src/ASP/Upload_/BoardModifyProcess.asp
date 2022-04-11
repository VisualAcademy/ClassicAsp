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
'[!] 작은따옴표 처리 : Name, Title, Content 정도만...
Name = Replace(Name, "'", "''")
Title = Replace(Title, "'", "''")
Content = Replace(Content, "'", "'')

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("ConnectionString"))
	strSql = "Select Password From Upload Where Password = '" & Password & "' And Num = " & Num
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	objRs.Open strSql, objCon
		If objRs.BOF Or objRs.EOF Then
			'패스워드 틀리다는 경고와 함께 뒤로 보내기
%>
			<script language="javascript">
			window.alert("비밀번호가 틀립니다.");
			window.history.go(-1);
			</script>
<%
		Else
			'업데이트 문 수행 후 View 페이지로 번호가지고 이동
			'변수처리 : 순수데이터 -> " & 변수 & "
			strSql = "Update Upload Set Name = '" & Name & "', Email = '" & Email & "', Title = '" & Title & "', Content = '" & Content & "', Encoding = '" & Encoding & "', Homepage = '" & Homepage & "', ModifyDate = GetDate(), ModifyIP = '" & Request.ServerVariables("REMOTE_HOST") & "' Where Num = " & Num & " And Password = '" & Password & "'"
			objCon.Execute(strSql)
			Response.Redirect("BoardView.asp?Num=" & Num)'//이동
		End If
	objRs.Close()
	Set objRs = Nothing
objCon.Close()
Set objCon = Nothing
%>










