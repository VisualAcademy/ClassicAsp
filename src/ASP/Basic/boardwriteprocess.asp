<%
'--------------------------------------------------
' Title : Basic 보드
' Program Name : boardwriteprocess.asp
' Program Description : 글쓰기 처리 페이지
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
'[1] 변수 선언
Option Explicit
Dim Name : Name = Request("Name")
Dim Email : Email = Request("Email")
Dim Homepage : Homepage = Request("Homepage")
Dim Title : Title = Request("Title")
Dim Content : Content = Request("Content")
Dim Encoding : Encoding = Request("Encoding")
Dim Password : Password = Request("Password")
Dim PostIP
	PostIP = Request.ServerVariables("REMOTE_HOST")
	
Dim strSql ' 홍길동 -> " & 변수 & "
strSql = "Insert Basic Values('" & Name & "', '" & Email & _
     "', '" & Title & "', GetDate(), '" & PostIP & _
      "', '" & Content & "', '" & Password & "', 0, '" & _
       Encoding & "', '" & Homepage & "', GetDate(), NULL)"
	
Dim objCon
'[2] 데이터베이스의 인스턴 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] 열기
objCon.Open(Application("ConnectionString"))
'[4] 명령 실행 : 홍길동 -> " & 변수 & "
objCon.Execute(strSql)
'[5] 닫기
objCon.Close()
'[6] 해제
Set objCon = Nothing
'[7] 리스트로 이동
Response.Redirect("./boardlist.asp")
%>