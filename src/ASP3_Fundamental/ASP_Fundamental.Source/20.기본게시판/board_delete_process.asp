<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_delete_process.asp
' Program Description : 전송된 패스워드가 맞으면 해당글 지우기
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
'변수의 명시적 선언
Option Explicit
Dim Num, Passwd, objDB, objRS, strSQL

Num = Request("Num")
Passwd = Request("Passwd")

'데이터베이스(connection객체)의 인스턴스 생성
Set objDB = Server.CreateObject("ADODB.Connection")

'데이터베이스 열기
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'레코드셋 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet")

'넘겨온 글번호에 해당되는 패스워드 가져오기
strSQL = "Select Passwd From my_board Where Num = " & Num
'Response.Write strSQL

'레코드 셋 열기
objRS.Open strSQL, objDB, 1


If objRS.Fields("Passwd") = Passwd Then
	'데이터베이스 처리(수정) 구문 작성
	strSQL = "Delete my_board Where Num =" & Num
        
	'데이터베이스 처리(입력) 실행
	objDB.Execute(strSQL)

	'방명록리스트 페이지로 리다이렉트 - 테스트할 때는 아래 코드는 주석처리.
	Response.Redirect "./board_list.asp"

Else
%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio.NET 7.0">
	</HEAD>
	<BODY>
		<script language="javascript">
			window.alert("비밀번호가 일치하지 않습니다");
			history.go(-1);
		</script>
	</BODY>
</HTML>
<%
End If
%>

<%
	'레코드셋 닫기
	objRS.Close

	'데이터베이스 닫기
	objDB.Close

	'레코드셋 개체의 인스턴스 해제
	Set objRS = Nothing

	'데이터베이스 객체의 인스턴스 해제
	Set objDB = Nothing
%>