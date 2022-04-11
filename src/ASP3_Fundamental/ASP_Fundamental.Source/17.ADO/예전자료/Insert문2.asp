<%
'[1] 변수 선언
Option Explicit
Dim objCon
'[2] 데이터베이스의 인스턴스 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] 데이터베이스 열기 : 데이터베이스연결문자열 지정
objCon.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")
'[4] SQL문 실행
objCon.Execute("Insert Memo Values('홍길동', 'hong@hong.com', '테스트입니다.', GetDate())")
'[5] 데이터베이스 닫기
objCon.Close()
'[6] 데이터베이스의 인스턴스 생성 해제
Set objCon = Nothing
%>