<%
'--------------------------------------------------
' Title : Myado Education Version
' Program Name : ado_admin_process.asp
' Program Description : ���±׸� �̿��� ������ �α���
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

   window.alert("�α��ε��� �ʾҽ��ϴ�\n�ٽ÷α��� �ϼ���.");
   history.go(-1);
   //history.back();
   
</script>
<%
End If 
%>

<html>
<head><title>������ �α�</title></head>
<body>
<center>���� �����ڷ� �������Դϴ�.</center>
<br>
<center>�ۺ��� ���������� �ش���� ��й�ȣ�� ��Ÿ���ϴ�.</center>
<br>
<center><a href="ado_list.asp">����Ʈ�� ���ư���</a></center>

</body>
</html>