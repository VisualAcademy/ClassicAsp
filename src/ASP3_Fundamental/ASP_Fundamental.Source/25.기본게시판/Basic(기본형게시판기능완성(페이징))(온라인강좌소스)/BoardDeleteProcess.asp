<%
Dim objCon: Dim objCmd: Dim objRs
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
	Set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = objCon
	'//넘겨져 온 번호와 비밀번호가 맞는 데이터가 있는지 검색
	objCmd.CommandText = "Select Password From Basic Where Num = " & Request("Num") & " And Password = '" & Request("Password") & "'"
	    Set objRs = objCmd.Execute()	
		If objRs.BOF Or objRs.EOF Then	'맞는 데이터가 없다면...
%>
		<SCRIPT LANGUAGE="JavaScript">
		window.alert("비밀번호가 맞지 않습니다.");
		history.go(-1);
		</SCRIPT>
<%
		Else	'맞는 데이터가 있다면 삭제 후 리스트로 이동
			objCmd.CommandText = "Delete Basic Where Num = " & Request("Num")
			objCmd.Execute()'//삭제 명령실행
			Response.Redirect("./BoardList.asp")'//리스트로 이동
		End If	
		objRs.Close()
		Set objRs = Nothing
	Set objCmd = Nothing
objCon.Close()
Set objCon = Nothing
%>