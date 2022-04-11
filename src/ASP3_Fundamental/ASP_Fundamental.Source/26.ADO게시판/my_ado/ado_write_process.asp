<%
'--------------------------------------------------------------------
' Title : 레코드셋 객체를 이용한 자료저장.
' Program Name : ado_write_process.asp
' Program Description : ADO 상수 사용법
' Include Files : ADO constants include file for VBScript
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>

<!-- HTML태그 및 SQL 특수문자  쓰기 제한 관련 사용자 정의함수 호출 -->
<!-- #include file="./ado_function.asp" -->

<!-- ------------------------------------------- -->
<!-- #include file="./scripts/adovbs.inc" -->
<!-- ------------------------------------------- -->

<%
'Option Explicit   '이 구문을 주석처리한 이유는 인클루드 파일에서 변수선언을 안했기때문(에러방지)

'변수 선언
Dim Name, Email, Title, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Content, objDB, objRS, strSQL

'Request객체의 Form컬렉션으로 넘어온 데이터(name, value)를 받는 구문.
Name = Request("Name")
Email = Request("Email")
Title = Request("Title")
PostDate = Date()
ModifyDate = Date() '처음 글쓸때도 해당 날짜 입력.
ReadCount = 0
PostIP = CStr(Request.ServerVariables("REMOTE_HOST"))
ModifyIP = CStr(Request.ServerVariables("REMOTE_HOST")) '처음 글쓸때도 해당 IP 입력.
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



'--------------------------------------------------------------------
' 레코드 셋 객체를 이용한 데이터 저장
'--------------------------------------------------------------------
'[1] 레코드 셋 객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet")

'[2]데이터베이스 연결 문자열 생성
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=ZZAZIPKIDOTCOM"

'[3] 레코드 셋 열기(여기서는 첫번째 인자값에 SQL문 대신에 해당 테이블 이름 지정)
objRS.Open "my_ado", strDB, adOpenStatic, adLockOptimistic, adCmdTable

	objRS.AddNew   '새로운 레코드를 생성할 공간 마련(Connection객체의 Insert문과 같음)

	objRS.Fields("Name") = Name
	objRS.Fields("Email") = Email
	objRS.Fields("Title") = Title
	objRS.Fields("PostDate") = Now()
	objRS.Fields("ModifyDate") = Now()
	objRS.Fields("ReadCount") = 0
	objRS.Fields("PostIP") = CStr(Request.ServerVariables("REMOTE_HOST"))
	objRS.Fields("ModifyIP") = CStr(Request.ServerVariables("REMOTE_HOST"))
	objRS.Fields("Encoding") = Encoding
	objRS.Fields("Passwd") = Passwd
	objRS.Fields("Content") = Content

	objRS.Update   '위에서 지정한 값을 삽입

	'레코드 셋 객체의 인스턴스 해제
	Set objRS = Nothing
'--------------------------------------------------------------------

'3. 게시판리스트 페이지로 리다이렉트 - 테스트할 때는 아래 코드는 주석처리.
Response.Redirect "./ado_list.asp"

%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<br>데이터 베이스가 입력되었습니다.
	</BODY>
</HTML>