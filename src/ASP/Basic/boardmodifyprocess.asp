<%
'--------------------------------------------------
' Title : Basic ����
' Program Name : boardmodifyprocess.asp
' Program Description : �����ϱ� ó�� : Update��
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
	Response.Write("�ش�Ǵ� �����Ͱ� �����ϴ�.")
Else
	If objRs("Password") = Password Then
	  ' ���������� -> " & ���� & "
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
	window.alert("��ȣ�� Ʋ���ϴ�.\n��ȣ�� Ȯ���ϼ���.");
	history.go(-1);
	</SCRIPT>
<%
	End If
End If	

objRs.Close(): objCon.Close()
Set objRs = Nothing: Set objCon = Nothing
%>