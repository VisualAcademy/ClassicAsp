<%
'--------------------------------------------------
' Title : Basic 보드
' Program Name : boarddeleteprocess.asp
' Program Description : 지우기 폼
' QueryString : delete.asp에서 Num=1, Password=1식으로 값이 전송됨...
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
Option Explicit
Dim Num: Num = Request("Num")
Dim Password: Password = Request("Password")
Dim objCon: Dim objRs: Dim strSql
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("CONNECTIONSTRING"))
Set objRs = Server.CreateObject("ADODB.RecordSet")
strSql = "Select Password From Basic" & _
		 " Where Num = "
strSql = strSql & Num
objRs.Open strSql, objCon
If objRs("Password") = Password Then '패스워드가 같다면,
	strSql = "Delete Basic Where Num = " & Num
	objCon.Execute(strSql)
	Response.Redirect("./boardlist.asp")
Else
%>
	<SCRIPT LANGUAGE="JavaScript">
	window.alert("암호가 틀립니다.\n암호를 확인하세요.");
	history.go(-1);
	</SCRIPT>
<%
End If
objRs.Close(): objCon.Close()
Set objRs = Nothing: Set objCon = Nothing
%>