<%
'--------------------------------------------------
' Title : Basic ����
' Program Name : boardwriteprocess.asp
' Program Description : �۾��� ó�� ������
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
'[1] ���� ����
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
	
Dim strSql ' ȫ�浿 -> " & ���� & "
strSql = "Insert Basic Values('" & Name & "', '" & Email & _
     "', '" & Title & "', GetDate(), '" & PostIP & _
      "', '" & Content & "', '" & Password & "', 0, '" & _
       Encoding & "', '" & Homepage & "', GetDate(), NULL)"
	
Dim objCon
'[2] �����ͺ��̽��� �ν��� ����
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] ����
objCon.Open(Application("ConnectionString"))
'[4] ��� ���� : ȫ�浿 -> " & ���� & "
objCon.Execute(strSql)
'[5] �ݱ�
objCon.Close()
'[6] ����
Set objCon = Nothing
'[7] ����Ʈ�� �̵�
Response.Redirect("./boardlist.asp")
%>