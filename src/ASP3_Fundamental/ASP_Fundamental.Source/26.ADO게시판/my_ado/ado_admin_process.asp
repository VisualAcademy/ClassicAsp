<%
'--------------------------------------------------
' Title : Myado Education Version
' Program Name : ado_admin_process.asp
' Program Description : 폼태그를 이용한 관리자 로그인
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
Option Explicit
Dim ID, Passwd

ID = LCase(Request("id"))
Passwd = Request("passwd")
  
If ID = "admin" AND Passwd = "1111" Then
	Session("ID") = "Admin"
Else
%>
<script language="javascript">

   window.alert("로그인되지 않았습니다\n다시로그인 하세요.");
   history.go(-1);
   //history.back();
   
</script>
<%
End If 
%>

<html>
<head><title>관리자 로긴</title></head>
<body>
<center>현재 관리자로 접속중입니다.</center>
<br>
<center>글보기 페이지에서 해당글의 비밀번호가 나타납니다.</center>
<br>
<center><a href="ado_list.asp">리스트로 돌아가기</a></center>

</body>
</html>