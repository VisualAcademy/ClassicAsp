<%
	'[1] 변수의 명시적 선언
	Option Explicit

	'[2] 변수 선언
	Dim objDB, strSQL
	Dim Name: Name="홍길동"
	Dim Email: Email="h@h.com"
	Dim Title: Title="홍길동입니다."
	Dim PostIP: PostIP="127.0.0.1"

	'[3] 데이터베이스의 인스턴스(실체)를 생성
	Set objDB = Server.CreateObject("ADODB.Connection")

	'[4] 데이터베이스 열기(Open() 메소드 사용 : DSN, UID, PWD)
	objDB.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")

	'[5] SQL 문자열 작성
	'홍길동 -> " & 변수 & "
	strSQL = "Insert Memo(Name, Email, Title, PostDate, PostIP) Values('" & Name & "','" & Email & "','" & Title & "',GetDate(),'" & PostIP & "')"

	'[6] SQL문 디버깅
	Response.Write strSQL

	'[7] SQL 문자열 실행(Execute()메소드 사용)
	objDB.Execute strSQL

	'[8] 데이터베이스 닫기
	objDB.Close

	'[9] 데이터베이스의 인스턴스 생성 해제
	Set objDB = Nothing
%>
<center>입력되었습니다.</center>
