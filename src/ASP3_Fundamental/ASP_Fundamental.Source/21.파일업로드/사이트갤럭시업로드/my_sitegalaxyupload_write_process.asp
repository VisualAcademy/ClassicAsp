<%
' -----------------------------------------------------------
' Title : Using SiteGalaxyUpload Component
' Program Name : my_sitegalaxyupload_write_process.asp
' Description : 파일올리기 처리 프로세스
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: 
' Support: 
' -----------------------------------------------------------
%>
<%
'변수의 명시적 선언
Option Explicit
Dim  objUploadForm, objFSO, FileName, strDirectory, AttachFile, FileSize
Dim strName, strExt, strFileName, tempFileName, i

'스크립팅 타임 아웃
Server.ScriptTimeout = 900    

'사이트갤럭시 업로드 컴포넌트의 인스턴스 생성
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form") 

'FSO 객체의 인스턴스 생성
Set objFSO = CreateObject("Scripting.FileSystemObject")

'파일명 추출
FileName = objUploadForm("FileName") 

'파일명에서 공백 제거
FileName = Trim(FileName)        

IF Len(FileName) > 0 Then '첨부한 파일이 있으면
'서버내 파일이 저장되는 위치지정 : 같은 경로의 files폴더, files폴더를 미리 만들어둬야 에러가 안남.
strDirectory = Server.MapPath(".") + "\files\" 
Response.Write "[1] 서버내 파일 저장 위치 : " & strDirectory & "<br>"

AttachFile = objUploadForm("FileName").FilePath    '저장될 파일의 경로 얻음(클라이언트 파일 위치)   
Response.Write "[2] 업로드 파일의 위치 : " & AttachFile & "<br>"

FileSize = objUploadForm("FileName").Size '파일 사이즈         
Response.Write "[3] 업로드 파일의 크기 : " & FileSize & "<br>"

FileName = Mid(AttachFile, InstrRev(AttachFile, "\")+1) '파일명.확장자 를 얻음!
Response.Write "[4] 파일명.확장자 : " & FileName & "<br>"
  
If InstrRev(FileName,".") Then
	strName = Mid(FileName, 1, InstrRev(FileName, ".")-1) '파일명 추출
Else
	strName = FileName
End If
Response.Write "[5] 파일명 : " & strName & "<br>"

strExt = Mid(FileName, Instr(FileName,".")+1) '확장자 추출
Response.Write "[6] 확장자 : " & strExt & "<br>"

'중복된 파일이 있으면 이름 뒤에 숫자를 붙임.
i = 0
Do While objFSO.FileExists(strDirectory & FileName)

	i = i + 1
	tempFileName = strName & "(" & i & ")"

	If strExt <> "" Then
		FileName = tempFileName & "." & strExt
	Else
		FileName = tempFileName
	End If

Loop

strFileName = strDirectory & FileName '서버에 저장할 경로와 파일 설정

Response.Write "[7] 서버에 저장할 경로와 파일명.확장자 : " & strFileName & "<br>"

objUploadForm("FileName").SaveAs strFileName '서버에 파일을 저장

End If 

Set objFSO = Nothing
Set objUploadForm = Nothing

Response.Write "[8] " & strFileName & "이 저장되었습니다."
%>
