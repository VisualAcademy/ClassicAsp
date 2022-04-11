<%
'--------------------------------------------------
' Title : Myupload Education Version
' Program Name : upload_modify_process.asp
' Program Description : 수정폼에서 전송된 값 Update문으로 저장
' Include Files : upload_function.asp
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!-- HTML태그 및 SQL 특수문자  쓰기 제한 관련 사용자 정의함수 호출 -->
<!-- #include file="./upload_function.asp" -->

<%
'Option Explicit   '이 구문을 주석처리한 이유는 인클루드 파일에서 변수선언을 안했기때문(에러방지)

'변수 선언
Dim Num, Name, Email, Title, ModifyDate, ModifyIP, Encoding, Passwd, Content, objDB, strSQL

'Request객체의 Form컬렉션으로 넘어온 데이터(name, value)를 받는 구문.
Num = Request.Form("Num")
Name = Request.Form("Name")
Email = Request.Form("Email")
Title = Request.Form("Title")
ModifyDate = Date() '처음 글쓸때도 해당 날짜 입력.
Passwd = Request.Form("password")
ModifyIP = Request.ServerVariables("REMOTE_HOST") '처음 글쓸때도 해당 IP 입력.
Encoding = Request.Form("Encoding")
Content = Request.Form("Content")

'HTML태그 및 SQL 특수문자  쓰기 제한.-----------------------------------
	Title = ConvertToHTML(Title)   '제목에 HTML사용금지
	Email = ConvertToHTML(Email)   'Email에 HTML사용금지
	
	Title = ConvertToSQL(Title)   'Title에 작은따옴표 처리
	Email = ConvertToSQL(Email)   'Email에 작은따옴표 처리
	Content = ConvertToSQL(Content)   'Content에 작은따옴표 처리

	'내용 중 인코딩이 HTML인지 Text인지에 따른 저장 형식 결정.
	Select Case Encoding

		Case "HTML"   'HTML을 처리해서 보여줄때
	
			Content = CStr(Content)
		
		Case "Text"   'HTML 소스를 그대로 보여줄 때

			Content = CStr(Content)
			Content = ConvertToHTML(Content)   

	End Select
'HTML태그 및 SQL 특수문자  쓰기 제한.-----------------------------------

'1. 데이터가 제대로 넘어왔는지 디버깅
'Response.Write(Name & "," & Email & "," & Title & "," & PostDate & "," & ModifyDate & "," & ReadCount & "," & PostIP & "," & ModifyIP & "," & Encoding & "," & Passwd& "," & Content)

'데이터베이스(connection객체)의 인스턴스 생성
Set objDB = Server.Createobject("ADODB.Connection")

'데이터베이스 열기
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'레코드셋 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet")

'넘겨온 글번호에 해당되는 패스워드 가져오기
strSQL = "Select Passwd From my_upload Where Num = " & Num
'Response.Write strSQL

'레코드 셋 열기
objRS.Open strSQL, objDB, 1

If objRS.Fields("Passwd") = Passwd Then
	'데이터베이스 처리(수정) 구문 작성
	strSQL2 = "Update my_upload Set Name = '" & Name & "', Email = '" & Email & "', Title = '" & Title & "', ModifyDate = '" & ModifyDate & "', Encoding = '" & Encoding & "', ModifyIP = '" & ModifyIP & "', Content = '" & Content & "' Where Num = " & Num
	'Response.Write strSQL2
        
	'데이터베이스 처리(입력) 실행
	objDB.Execute(strSQL2)

	'레코드셋 닫기
	objRS.Close

	'데이터베이스 닫기
	objDB.Close

	'레코드셋 개체의 인스턴스 해제
	Set objRS = Nothing

	'데이터베이스 객체의 인스턴스 해제
	Set objDB = Nothing

	'방명록리스트 페이지로 리다이렉트 - 테스트할 때는 아래 코드는 주석처리.
	Response.Redirect "./upload_list.asp"

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
