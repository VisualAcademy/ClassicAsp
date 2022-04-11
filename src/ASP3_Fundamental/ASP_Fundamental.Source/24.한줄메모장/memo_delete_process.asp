<%
'--------------------------------------------------------------------
' MyMemo Education Version
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<%
Option Explicit

'변수 선언
Dim objDB, strSQL, Num, Passwd

Num = Request.Form("Num")
Passwd = Request.Form("passwd")

If Num <> "" AND Passwd = "1234" Then
	'데이터베이스(connection객체)의 인스턴스 생성
	Set objDB = Server.CreateObject("ADODB.Connection")

	'데이터베이스 열기
	objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

	'데이터베이스 처리(삭제) 구문 작성
	strSQL = "Delete my_memo Where Num =" & Num
        
	'데이터베이스 처리(입력) 실행
	objDB.Execute(strSQL)

	'메모리스트 페이지로 리다이렉트 - 테스트할 때는 아래 코드는 주석처리.
	Response.Redirect "./memo_list.asp"

Else
%>
<SCRIPT language="javascript">
	history.go(-1);
</SCRIPT>
<%
End If
%>
<%
	'데이터베이스 닫기
	objDB.Close

	'데이터베이스 객체의 인스턴스 해제
	Set objDB = Nothing
%>
