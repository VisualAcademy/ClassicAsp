<%
'--------------------------------------------------
' Title : Basic 보드
' Program Name : boardmodifyprocess.asp
' Program Description : 수정하기 처리 : Update문
' QueryString : modify.asp -> Num, Name, ..., Password
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
Option Explicit
Dim Num: Num = Request("Num")'//
Dim Name: Name = Request("Name")
Dim Email: Email = Request("Email")
Dim Title: Title = Request("Title")
Dim Content: Content = Request("Content")
Dim Homepage: Homepage = Request("Homepage")
Dim Encoding: Encoding = Request("Encoding")
Dim Password: Password = Request("Password")
Dim objCon: Dim objRs: Dim strSql
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("ConnectionString"))
Set objRs = Server.CreateObject("ADODB.RecordSet")
strSql = "Select Password From Basic Where Num = " _
		 & Num
objRs.Open strSql, objCon 

If objRs.BOF Or objRs.EOF Then
	Response.Write("해당되는 데이터가 없습니다.")
Else
	If objRs("Password") = Password Then
	  ' 순수데이터 -> " & 변수 & "
    strSql = "Update Basic Set Name = '" & Name & _
     "', Email = '" & Email & "', Title = '" & Title & _
      "', Content = '" & Content & "', Encoding = '" & _
       Encoding & "', Homepage = '" & Homepage & _
        "', ModifyDate = GetDate(), ModifyIP = '" & _
         Request.ServerVariables("REMOTE_ADDR") & _
          "' Where Num = " & Num
    'Response.Write(strSql)
    'Response.End()                
		objCon.Execute(strSql)
		Response.Redirect("./boardview.asp?Num=" & Num)
	Else
%>
	<SCRIPT LANGUAGE="JavaScript">
	window.alert("암호가 틀립니다.\n암호를 확인하세요.");
	history.go(-1);
	</SCRIPT>
<%
	End If
End If	

objRs.Close(): objCon.Close()
Set objRs = Nothing: Set objCon = Nothing
%>