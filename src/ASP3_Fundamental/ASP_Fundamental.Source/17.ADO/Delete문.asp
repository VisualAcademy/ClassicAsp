<%
'[1] 변수 선언
Option Explicit
Dim objCon
Dim strSql
'[2] 데이터베이스의 인스턴스 생성 : Connection 오브젝트
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] 데이터베이스 열기
objCon.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")

'[4] SQL문 실행
strSql = "Delete Memo Where Num = 5"
objCon.Execute(strSql)

'[5] 데이터베이스 닫기
objCon.Close()
'[6] 데이터베이스의 인스턴스 해제
Set objCon = Nothing 
%>