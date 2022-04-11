<%
'1. 사이트갤럭시 업로드 컴포넌트 설치
'2. 글쓰기 폼 페이지에서...
'	<form> 태그에
'	enctype="multipart/form-data" 속성 추가
'3. 글쓰기 처리 페이지에서... Request객체를 사용 못하므로
'	사이트갤럭시 업로드 컴포넌트의 인스턴스 생성
'4. 각각의 필드를 인스턴스 이름으로 받기
'5. 넘겨져 온 전체 경로 값 받기
'6. 넘겨져 온 파일의 크기 값 받기
'7. 전체 경로에서 파일명 분리
'8. 파일명에서 순수파일명과 확장자 분리
'9. 업로드할 폴더 정의

'자료실 3 : 사이트 갤럭시 업로드 컴포넌트의 인스턴스 생성-----------------
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form")
'---------------------------------------------------------------

'자료실 4 : Request객체 대신에 objUploadForm객체로 변경---------------
'변수 선언
Dim objCon: Dim strSql
Dim Name: Name = objUploadForm("Name")
Dim Email: Email = objUploadForm("Email")
Dim Title: Title = objUploadForm("Title")
Dim PostIP: PostIP = Request.ServerVariables("REMOTE_HOST")
Dim Content: Content = objUploadForm("Content")
Dim Password: Password = objUploadForm("Password")
Dim Encoding: Encoding = objUploadForm("Encoding")
Dim Homepage: Homepage = objUploadForm("Homepage")
'Dim FileName: FileName = objUploadForm("FileName")
'Dim FileSize: FileSize = 100
'---------------------------------------------------------------

'자료실 5 : 앞에서 넘겨져 온 파일의 전체 경로값 받기---------------------
Dim FullFilePath
FullFilePath = objUploadForm("FileName").FilePath
'---------------------------------------------------------------

'자료실 6 : 앞에서 넘겨져 온 파일의 크기(바이트 단위)--------------------
Dim FileSize
FileSize = objUploadForm("FileName").Size
'---------------------------------------------------------------

'자료실 7 : 전체 경로에서 파일명과 확장자 분리--------------------
Dim FileName
FileName = Mid(FullFilePath, InStrRev(FullFilePath, "\") + 1)
'---------------------------------------------------------------

'자료실 8 : 파일명에서 순수 파일명과 확장자 따로 분리-------------------
Dim strName	'순수 파일명
strName = Left(FileName, InStrRev(FileName, ".") - 1)
Dim strExt	'순수 확장자
strExt = Mid(FileName, InStrRev(FileName, ".") + 1)
'---------------------------------------------------------------

'자료실 9 : 어디로 업로드 할 건지 폴더명 지정-------------------
Dim strDir
strDir = Server.MapPath(".") & "\files\"  '반드시 쓰기 권한 필요
'---------------------------------------------------------------

'자료실 10 : 파일명 중복처리-------------------
Dim objFso
Set objFso = Server.CreateObject("Scripting.FileSystemObject")
i = 0
Do While objFSO.FileExists(strDir + FileName)
	i = i + 1
	tempFileName = strName & "(" & i & ")"
	If strExt <> "" Then
		FileName = tempFileName & "." & strExt
	Else
		FileName = tempFileName
	End If
Loop
'---------------------------------------------------------------

'자료실 11 : 최종 저장-------------------
objUploadForm("FileName").SaveAs strDir & FileName
'---------------------------------------------------------------

'[!] 작은따옴표 처리
Name = Replace(Name, "'", "''")
Title = Replace(Title, "'", "''")
Content = Replace(Content, "'", "''")
'커넥션 객체 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'커넥션 객체 오픈
objCon.Open(Application("ConnectionString"))
'미리 실행할 SQL문 지정 : 순수한데이터 -> " & 변수 & "
strSql = "Insert Upload(Name,Email,Title,PostDate,PostIP,Content,Password,ReadCount,Encoding,Homepage,ModifyDate,ModifyIP,FileName,FileSize,DownCount) Values('" & Name & "','" & Email & "','" & Title & "',GetDate(),'" & PostIP & "','" & Content & "',	'" & Password & "',0,'" & Encoding & "','" & Homepage & "',Null,Null,'" & FileName & "'," & FileSize & ",0)"
'SQL문 실행
objCon.Execute(strSql)
'커넥션 객체 닫기
objCon.Close()
'커넥션 객체 해제
Set objCon = Nothing
'리스트 페이지로 이동
Response.Redirect("BoardList.asp")
%>
<center>데이터가 입력되었습니다.</center>