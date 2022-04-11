<%
Dim FileName '//넘겨져 온 파일명
Dim strDir '//서버측 저장 경로
Dim strName '//순수한 파일명 : test.gif -> test
Dim strExt '//순수한 확장자 : test.gif -> gif
'[0]사이트갤럭시 업로드 컴포넌트의 인스턴스 생성
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form")
'[1]넘어온 파일의 전체 경로 값 출력 : FilePath 속성
Response.Write("클라이언트 전체 경로 : " & objUploadForm("FileName").FilePath & "<br>")
'[2]넘어온 파일의 크기(바이트 단위) : Size 속성
Response.Write("크기 : " & objUploadForm("FileName").Size) 
FileName = objUploadForm("FileName")'//클라이언트 경로값 받기:FilePath
FileName = Mid(FileName, InStrRev(FileName, "\") + 1)'//파일명분리
strName = Left(FileName, InStrRev(FileName, ".") - 1)'파일명구하기
strExt = Mid(FileName, InStrRev(FileName, ".") + 1)'확장자구하기
'[!]파일명 중복 처리 : 파일명 뒤에 순서대로 번호를 매김
Dim objFso
Set objFso = Server.CreateObject("Scripting.FileSystemObject")
'파일명 뒤에 번호 붙이기
i = 0
Do While objFSO.FileExists(Server.MapPath(".") & "\files\" + FileName)
	i = i + 1
	tempFileName = strName & "(" & i & ")"
	If strExt <> "" Then
		FileName = tempFileName & "." & strExt
	Else
		FileName = tempFileName
	End If
Loop
'[!]서버측 저장 경로와 파일명
strDir = Server.MapPath(".") & "\files\" & FileName	'//
'[3]업로드하기 : SaveAs() 메서드 사용
objUploadForm("FileName").SaveAs strDir				'//
%>












