<%
'--------------------------------------------------
' Title : Myado Education Version
' Program Name : ado_delete_process.asp
' Program Description : 전송된 패스워드가 맞으면 해당글 지우기
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!-- #include file="./scripts/adovbs.inc" -->

<%
'변수의 명시적 선언
'Option Explicit
Dim Num
Dim Passwd
Dim objRS
Dim strDB
Dim strSQL

Num = Request("Num")
Passwd = Request("Passwd")

'[1] 레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'[2] 데이터베이스 연결 문자열 생성
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=XPPLUSDOTNET"

'[3] 넘겨온 글번호에 해당되는 패스워드 가져오기
strSQL = "Select * From my_ado Where Num = " & Num
'Response.Write strSQL

'[4] 레코드셋 열기
objRs.Open strSQL, strDB, adOpenStatic, adLockPessimistic, adCmdText
'objRS.Open strSQL, Application("Connection_string"), adOpen...........   <--- global.asa에 미리 애플리케이션 변수로 선언했을 경우.

If objRS.Fields("Passwd") = Passwd Then
	'데이터베이스 처리(수정) 구문 작성

	objRS.Delete   '현재 불려진 레코드셋 삭제

	objRS.Update   '삭제 진행.

	'레코드셋 닫기
	objRS.Close

	'레코드셋 개체의 인스턴스 해제
	Set objRS = Nothing

	'방명록리스트 페이지로 리다이렉트 - 테스트할 때는 아래 코드는 주석처리.
	Response.Redirect "./ado_list.asp"

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

