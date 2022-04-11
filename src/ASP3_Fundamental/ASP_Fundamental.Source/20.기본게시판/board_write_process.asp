<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_write_process.asp
' Program Description : 입력폼에서 전송된 값 Insert문으로 저장
' Include Files : board_function.asp
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!-- HTML태그 및 SQL 특수문자  쓰기 제한 관련 사용자 정의함수 호출 -->
<!-- #include file="./board_function.asp" -->

<%
'Option Explicit   '이 구문을 주석처리한 이유는 인클루드 파일에서 변수선언을 안했기때문(에러방지)

'변수 선언
Dim Name, Email, Title, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Content, objDB, strSQL

'Request객체의 Form컬렉션으로 넘어온 데이터(name, value)를 받는 구문.
Name = Request.Form("Name")
Email = Request.Form("Email")
Title = Request.Form("Title")
PostDate = Date()
ModifyDate = Date() '처음 글쓸때도 해당 날짜 입력.
ReadCount = 0
PostIP = Request.ServerVariables("REMOTE_HOST")
ModifyIP = Request.ServerVariables("REMOTE_HOST") '처음 글쓸때도 해당 IP 입력.
Encoding = Request.Form("Encoding")
Passwd = Request.Form("Password")
Content = Request.Form("Content")

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

'1. 데이터가 제대로 넘어왔는지 디버깅
'Response.Write(Name & "," & Email & "," & Title & "," & PostDate & "," & ModifyDate & "," & ReadCount & "," & PostIP & "," & ModifyIP & "," & Encoding & "," & Passwd& "," & Content)

'데이터베이스(connection객체)의 인스턴스 생성
Set objDB = Server.Createobject("ADODB.Connection")

'데이터베이스 열기
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'데이터베이스 처리(입력) 구문 작성
strSQL = "Insert my_board(Name, Email, Title, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Content) Values " & _
         "('" & Name & "','" & Email & "','" & Title & "','" & PostDate & "','" & _
         ModifyDate & "'," & ReadCount & ",'" & PostIP & "','" & ModifyIP & "','" & _
		 Encoding & "','" & Passwd & "','" & Content & "')"
        
'2. 쿼리 구문 디버깅
'Response.Write(strSQL)

'데이터베이스 처리(입력) 실행
objDB.Execute(strSQL)

'데이터베이스 닫기
objDB.Close

'데이터베이스 객체의 인스턴스 해제
Set objDB = Nothing

'3. 게시판리스트 페이지로 리다이렉트 - 테스트할 때는 아래 코드는 주석처리.
Response.Redirect "./board_list.asp"

%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<br>데이터 베이스가 입력되었습니다.
	</BODY>
</HTML>