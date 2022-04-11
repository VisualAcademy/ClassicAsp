<%
'--------------------------------------------------------------------
' Title : MyMemo Education Version
' Program Name : memo_write_process.asp
' Program Description : 메모글 저장하는 프로세스
' Include Files : None
' Last Update : 2001.08.26. 02:20
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<%
'[1] 변수의 명시적 선언
Option Explicit

'[2] 변수 선언
Dim Name	'memo_write_form에서 전송된 이름값을 저장할 변수
Dim Email	'memo_write_form에서 전송된 이메일값을 저장할 변수
Dim Title		'memo_write_form에서 전송된 제목(메모)값을 저장할 변수
Dim PostDate	'MY메모장 DB에 기록된 날짜를 저장할 변수
Dim objDB		'Connection객체의 인스턴스 객체를 저장할 변수
Dim strSQL		'MY메모장 DB에 데이터를 저장하는 SQL문을 저장하는 변수 

'[3] 변수 값 대입 : Request객체의 Form컬렉션으로 넘어온 데이터(name, value)를 받는 구문.
Name = Request.Form("name")
Email = Request.Form("email")
Title = Request.Form("title")
PostDate = Date()   'Now()

'[4] 변수 값 처리 : HTML태그 및 SQL 특수문자 쓰기 제한(DB 저장 에러 방지)
Title = Replace(Title,"'","''")   '작은(홑)따옴표 처리
Title = Replace(Title,"&","&amp;")   '&(앰퍼샌드) 기호 처리
Title = Replace(Title,"<","&lt;")   '< 기호 처리
Title = Replace(Title,">","&gt;")   '> 기호 처리

'[5] 데이터가 제대로 넘어왔는지 디버깅(항상 디버깅 하는 습관을 들여놓을 것)
'Response.Write(Name & "," & Email & "," & Title & "," & PostDate)

'[6] 데이터베이스(Connection객체)의 인스턴스 생성
Set objDB = Server.CreateObject("ADODB.Connection")

'[7] 데이터베이스 열기(ODBC설정의 DSN정보와 SQL Server의 로그인 정보(uid/pwd) 사용)
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'[8] 데이터베이스 처리(입력) 구문 작성(작은 따옴표 사용 주의)
'ASP 날짜 함수 사용 입력, "& _" 연산자로 연결
'strSQL = "Insert my_memo(Name, Email, Title, PostDate) Values " & _
'"('" & Name & "','" & Email & "','" & Title & "'," & PostDate & ")"
'SQL 날짜 함수 사용 입력, "&" 누적 연산자로 연결
strSQL = "Insert my_memo(Name, Email, Title, PostDate) Values "
strSQL = strSQL & "('" & Name & "','" & Email & "','" & Title & "',GetDate())"
'----------------------------------------------------------------------------------------
'Insert my_memo Values ('박용준','abc@a.com','메모''이져?','2001-11-13')
'Insert my_memo Values ('   박용준  ','   abc@a.com   ','   메모''이져?  ',' 2001-11-13 ')
'"Insert my_memo Values ('"  &  박용준 & "','"  & abc@a.com &  "','"  & 메모''이져? & "','" 2001-11-13 & "')"
'strSQL = "Insert my_memo Values ('"  &  박용준 & "','"  & abc@a.com &  "','"  & 메모''이져? & "','" 2001-11-13 & "')"
'strSQL = "Insert my_memo Values ('"  & Name & "','"  & Email &  "','"  & Title & "','" PostDate & "')"
'----------------------------------------------------------------------------------------

'[9] 쿼리 구문 디버깅 : 아래 출력결과를 SQL Server 쿼리분석기에서 실행시켜본다.
'Response.Write(strSQL)

'[10] 데이터베이스 처리(입력) 실행
objDB.Execute(strSQL)

'[11] 데이터베이스 닫기
objDB.Close

'[12] 데이터베이스 객체의 인스턴스 해제
Set objDB = Nothing

'[13] 메모리스트 페이지로 리다이렉트 - 테스트할 때는 아래 코드는 주석처리.
Response.Redirect "./memo_list.asp"
%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<CENTER>
			데이터 베이스가 입력되었습니다.
		</CENTER>
	</BODY>
</HTML>
