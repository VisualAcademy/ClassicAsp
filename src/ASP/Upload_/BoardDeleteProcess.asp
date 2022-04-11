<%
'변수 선언 및 쿼리스트링 받기
Dim objCon: Dim strSql: Dim objRs
Dim Num: Num = Request("Num")
Dim Password: Password = Request("Password")
'커넥션 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'연결
objCon.Open(Application("ConnectionString"))
	'SQL문 지정 : 넘겨져 온 번호에 해당하는 패스워드/번호삭제 
	strSql = "Select * From Upload Where Num = " & Num & _
			 " And Password = '" & Password & "'"
	'레코드셋 객체로 받기 : 생성과 동시에 명령 실행
	Set objRs = objCon.Execute(strSql)'//objRs.Open()
	If objRs.BOF Or objRS.EOF Then
		'뒤로 보내기 : 자바스크립트 사용
%>
		<script language="javascript">
		window.alert("해킹하려구?");
		window.history.go(-1);
		</script>
<%		
		'Response.End()
	Else
	
'-----------------------------------------------
'이미 업로드된 파일을 먼저 삭제하는 로직
'-----------------------------------------------
Call DeleteArticleNow()
Sub DeleteArticleNow()
	Dim objFSO, objFile	'//FSO객체용 변수 선언
	If objRS.Fields("FileName") <> "" Then '//업로드된 파일이 있다면
		'//FSO객체의 인스턴스 생성
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		'//같은경로의 files폴더에있는 파일에 대한 정보를 얻어와서...
		Set objFile = objFSO.GetFile(Server.MapPath(".") + "\files\" & objRS.Fields("FileName"))
		'//해당 파일을 삭제
		objFile.Delete
		'//객체 해제
		Set objFile = Nothing
		Set objFSO = Nothing
	End If
End Sub
'----------------------------------------------

		'삭제 처리 후 리스트로 이동
		strSql = "Delete Upload Where Num = " & Num
		objCon.Execute(strSql)
		Response.Redirect("BoardList.asp")
	End If
	'레코드셋(메모리상의 테이블 결과) 닫기
	objRs.Close()
	'레코드셋 객체 해제
	Set objRs = Nothing
'닫기
objCon.Close()
'커넥션 해제
Set objCon = Nothing
'Response.Redirect("BoardList.asp")
%>








