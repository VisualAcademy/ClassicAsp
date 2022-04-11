<%
' -----------------------------------------------------------
' Title : Using SiteGalaxyUpload Component
' Program Name : my_sitegalaxyupload_write_process.asp
' Description : 파일올리기 처리 프로세스(중복 파일 1개 처리)
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
' -----------------------------------------------------------
%>
<%
'변수의 명시적 선언
Option Explicit
Dim  objUploadForm, objFSO, FileName, strDirectory, AttachFile, FileSize, strName, strExt, strFileName, countFileName, bExist

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
strDirectory = Server.MapPath(".") + "\file\" '서버내 파일이 저장되는 위치지정
Response.Write "[1] 서버내 파일 저장 위치 : " & strDirectory & "<br>"

AttachFile = objUploadForm("FileName").FilePath    '저장될 파일의 경로 얻음(클라이언트 파일 위치)   
Response.Write "[2] 업로드 파일의 위치 : " & AttachFile & "<br>"

FileSize = objUploadForm("FileName").Size '파일 사이즈         
Response.Write "[3] 업로드 파일의 크기 : " & FileSize & "<br>"

FileName = Mid(AttachFile, InstrRev(AttachFile, "\")+1) '파일명.확장자 를 얻음!
Response.Write "[4] 파일명.확장자 : " & FileName & "<br>"
  
strName = Mid(FileName, 1, Instr(FileName, ".")-1) '파일명 추출
Response.Write "[5] 파일명 : " & strName & "<br>"

strExt = Mid(FileName, Instr(FileName,".")+1) '확장자 추출
Response.Write "[6] 확장자 : " & strExt & "<br>"

bExist = True '같은 파일이 존재한다고 가정
strFileName = strDirectory&strName&"."&strExt '서버에 저장할 경로와 파일 설정
Response.Write "[7] 서버에 저장할 경로와 파일명.확장자 : " & strFileName & "<br>"



countFileName = 0 '같은 파일일때 파일명뒤에 붙일 숫자
        
IF bExist = True Then
	IF (objFSO.FileExists(strFileName)) Then '같은 파일이 존재하면
		countFileName = countFileName+1
		FileName = strName&"_"&countFileName&"."&strExt
		strFileName = strDirectory&FileName '실제 서버에 저장될 경로와 파일명
	Else
		bExist = False
	End IF
End IF
Response.Write "[8] 실제서버에 저장될 파일명 : " & strFileName & "<br>"

objUploadForm("FileName").SaveAs strFileName '서버에 파일을 저장

End If 

Set objFSO = Nothing
Set objUploadForm = Nothing

Response.Write "<br>" & strFileName & "이 저장되었습니다."
%>
