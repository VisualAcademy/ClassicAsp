<%
	'[1] 변수의 명시적 선언
	Option Explicit

	'[2] 변수 선언
	Dim objDB, strSQL

	'[3] 데이터베이스의 인스턴스(실체)를 생성
	Set objDB = Server.CreateObject("ADODB.Connection")
	
	'[4] 데이터베이스 열기(Open() 메소드 사용 : DSN, UID, PWD)
	objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")
	
	'[5] SQL 연결 문자열 작성
	strSQL = "Update my_memo Set Email = 'bbb@bbb.com'"

	'[6] SQL 문자열 실행(Execute()메소드 사용)
	objDB.Execute strSQL

	'[7] 데이터베이스 닫기
	objDB.Close

	'[8] 데이터베이스의 인스턴스 생성 해제
	Set objDB = Nothing
%>
<center>수정되었습니다.</center>
