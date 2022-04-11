<%
Dim strSql
strSql = "Insert Memos(Name, Email, Title, PostIP) " & _
    " Values('홍길동', 'h@h.com', '안녕', '127.0.01') "
Dim objCon

'[1] 데이터베이스 연결을 위한 커넥션 객체 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'[2] 데이터베이스 오픈(연결) : 데이터베이스 연결 문자열 필요
objCon.Open("Provider=SQLOLEDB.1;Password=MemoPwd;Persist Security Info=True;User ID=MemoUid;Initial Catalog=MemoDB;Data Source=.\SQLExpress")
'[3] 명령어 실행
objCon.Execute(strSql)
'[4] 데이터베이스 닫기
objCon.Close()
'[5] 데이터베이스 객체 생성 해제
Set objCon = Nothing
%>
데이터가 저장되었습니다.