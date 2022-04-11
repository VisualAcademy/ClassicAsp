<%
	'[1] 변수의 명시적 선언
	Option Explicit

	'[2] 변수 선언
	Dim objDB, objRS, strSQL

	'[3] 데이터베이스의 인스턴스(실체)를 생성
	Set objDB = Server.CreateObject("ADODB.Connection")

	'[4] 데이터베이스 열기(Open() 메소드 사용 : DSN, UID, PWD)
	objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

	'[5] 레코드셋의 인스턴스 생성
	Set objRS = Server.CreateObject("ADODB.RecordSet")

	'[6] SQL문 작성
	strSQL = "Select * From my_memo Order By Num Desc"

	'[7] 레코드셋 열기(Open()메소드 : SQL문, DB연결자(커넥션객체), 커서타입, 락타입, 옵션)
	objRS.Open strSQL, objDB

	'[8] 레코드 값 처리 프로세스 영역
	If objRS.BOF OR objRS.EOF Then
		Response.Write("<center>입력된 자료가 없습니다.</center>")
	Else
		Response.Write("<div align='center'>")
		Response.Write("<table border='1'>")
		Response.Write("<tr><td>번호</td><td>이름</td><td>이메일</td><td>제목</td><td>작성일</td></tr>")
		
		Do Until objRS.EOF
			Response.Write("<tr>")
			Response.Write("<td>" & objRS.Fields("Num") & "</td>")
			Response.Write("<td>" & objRS.Fields("Name") & "</td>")
			Response.Write("<td>" & objRS.Fields("Email") & "</td>")
			Response.Write("<td>" & objRS.Fields("Title") & "</td>")
			Response.Write("<td>" & objRS.Fields("PostDate") & "</td>")
			Response.Write("</tr>")
			objRS.MoveNext	
		Loop

		Response.Write("</table>")
		Response.Write("</div>")
	End If

	'[9] 레코드셋 닫기
	objRS.Close

	'[10] 데이터베이스 닫기
	objDB.Close

	'[11] 레코드셋 개체의 인스턴스 해제
	Set objRS = Nothing

	'[12] 데이터베이스 개체의 인스턴스 해제
	Set objDB = Nothing
%>


