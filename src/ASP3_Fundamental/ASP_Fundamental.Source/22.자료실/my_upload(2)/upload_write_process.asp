<%
'--------------------------------------------------
' Title : MyUpload Education Version
' Program Name : upload_write_process.asp
' Program Description : 입력폼에서 전송된 값 Insert문으로 저장
' Include Files : upload_function.asp
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!-- HTML태그 및 SQL 특수문자  쓰기 제한 관련 사용자 정의함수 호출 -->
<!-- #include file="./upload_function.asp" -->

<%
'Option Explicit   '이 구문을 주석처리한 이유는 인클루드 파일에서 변수선언을 안했기때문(에러방지)

'변수 선언
Dim Name, Email, Title, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Content, objDB, strSQL
Dim Num, Ref, Step, RefOrder, Number, objRS

' -------------------- 자료실에서의 추가 --------------------
Dim  objUploadForm, objFSO, FileName, strDirectory, AttachFile, FileSize, strSQL2
Dim strName, strExt, strFileName, tempFileName, i

'스크립팅 타임 아웃
Server.ScriptTimeout = 900    

'사이트갤럭시 업로드 컴포넌트의 인스턴스 생성
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form") 
' -------------------- 자료실에서의 추가 --------------------

'사이트갤럭시 업로드 컴포넌트의 객체를 사용해서 넘어온 데이터(name, value)를 받는 구문.
Num = objUploadForm("Num")
Name = objUploadForm("Name")
Email = objUploadForm("Email")
Title = objUploadForm("Title")
PostDate = Date()
ModifyDate = Date() '처음 글쓸때도 해당 날짜 입력.
ReadCount = 0
PostIP = Request.ServerVariables("REMOTE_HOST")
ModifyIP = Request.ServerVariables("REMOTE_HOST") '처음 글쓸때도 해당 IP 입력.
Encoding = objUploadForm("Encoding")
Passwd = objUploadForm("Password")
Content = objUploadForm("Content")

'HTML태그 및 SQL 특수문자  쓰기 제한.-----------------------------------
	Title = ConvertToHTML(Title)   '제목에 HTML사용금지
	Email = ConvertToHTML(Email)   'Email에 HTML사용금지
	
	Title = ConvertToSQL(Title)   'Title에 작은따옴표 처리
	Email = ConvertToSQL(Email)   'Email에 작은따옴표 처리
	Content = ConvertToSQL(Content)   'Content에 작은따옴표 처리

	'내용 중 인코딩이 HTML인지 Text인지에 따른 저장 형식 결정.
	Select Case Encoding

		Case "HTML"   'HTML을 처리해서 보여줄때
	
			Content = CStr(Content)
		
		Case "Text"   'HTML 소스를 그대로 보여줄 때

			Content = CStr(Content)
			Content = ConvertToHTML(Content)   

	End Select
'HTML태그 및 SQL 특수문자  쓰기 제한.-----------------------------------

'데이터가 제대로 넘어왔는지 디버깅
Response.Write(Name & "," & Email & "," & Title & "," & PostDate & "," & ModifyDate & "," & ReadCount & "," & PostIP & "," & ModifyIP & "," & Encoding & "," & Passwd& "," & Content) & "<br>"


' -------------------- Q & A 게시판에서의 추가 --------------------
'데이터베이스(connection객체)의 인스턴스 생성
Set objDB = Server.Createobject("ADODB.Connection")

'데이터베이스 열기
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")
 
 '게시판의 총 레코드 수를 Number라는 이름으로 반환
strSQL = "Select Count(Num) As Number From my_upload"

'레코드셋 객체의 인스턴스 생성
Set objRS = Server.Createobject("ADODB.RecordSet")

'레코드셋의 인수는 무시할 것...(나중에 RecordSet객체 시간에 자세히 다룸.)
objRS.Open strSQL, objDB, 2, 1, 1

'If objRS.BOF OR objRS.EOF Then
'	Response.Write "데이터가 없슴다."
'End IF

If objRS.BOF OR objRS.EOF Then '괄호 주의
	Number = 1
Else 
	Number = objRS("Number") + 1
End If

'Response.Write Number
'Response.Write "Ref :" & Ref & "<br>"

If Num <> "" Then '즉 답변쓰기라면

	Ref = objUploadForm("Ref")
	Step = objUploadForm("Step")
	RefOrder = objUploadForm("RefOrder")

	strSQL = "Update my_upload Set RefOrder = RefOrder + 1 Where Ref = " & Ref & " AND RefOrder > " & RefOrder
	Response.Write strSQl & "<br>"


	objDB.Execute(strSQL)

	Step = Step + 1
	RefOrder = RefOrder + 1

Else ' 첨 글쓰기라면...

Ref = Number
Step = 0
RefOrder = 0

End If
' -------------------- Q & A 게시판에서의 추가 --------------------
'Response.End
' -------------------- 자료실에서의 추가 --------------------

'FSO 객체의 인스턴스 생성
Set objFSO = CreateObject("Scripting.FileSystemObject")

'파일명 추출
FileName = objUploadForm("FileName") 

If FileName = "" Then
	FileName = Null
	FileSize = 0
Else

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

strExt = Mid(FileName, InstrRev(FileName,".")+1) '확장자 추출
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



Response.Write "[8] " & strFileName & "이 저장되었습니다."
' -------------------- 자료실에서의 추가 --------------------

End If

'데이터베이스 처리(입력) 구문 작성 : 파일명과 파일크기를 데이터베이스에 저장
strSQL2 = "Insert my_upload(Name, Email, Title, FileName, FileSize, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Ref, Step, RefOrder, Content) Values "
strSQL2 = strSQL2 & "('" & Name & "','" & Email & "','" & Title & "','" & FileName & "'," 
strSQL2 = strSQL2 & FileSize & ",'" & PostDate & "','" 
strSQL2 = strSQL2 & ModifyDate & "'," & ReadCount & ",'" & PostIP & "','" & ModifyIP & "','"
strSQL2 = strSQL2 & Encoding & "','" & Passwd & "'," & Ref & "," & Step & "," & RefOrder & ",'" & Content & "')"
        
'쿼리 구문 디버깅
Response.Write strSQL2 & "<br>"

'데이터베이스 처리(입력) 실행
objDB.Execute(strSQL2)

%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<center>
		<br>데이터 베이스가 입력되었습니다. <br>
		<a href="./upload_list.asp">메인 리스트로 이동</a>
		</center>
	</BODY>
</HTML>

<%
'데이터베이스 닫기
objDB.Close
'데이터베이스 객체의 인스턴스 해제 : 컴포넌트 객체 해제는 반드시 모든 처리가 끝난후 한다.
Set objFSO = Nothing
Set objUploadForm = Nothing
Set objDB = Nothing
%>
