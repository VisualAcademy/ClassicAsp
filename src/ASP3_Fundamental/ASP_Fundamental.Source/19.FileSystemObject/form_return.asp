<html>
<body bgColor=#ffffff>

<% If Request("userid") = "" then %>
<p> �� ������ �Է� </p>
<hr>

<form action="form_return.asp" method="post"> 
   ����� ID : <input type="text" name="userid" size=20><br>
   ��й�ȣ : <input type="password" name="password" size=20><p>
   <input type="submit" value="������">
   <input type="reset" value="�ʱ�ȭ">
</form>

<% Else %>
<p> �Է��� ��� </p>
<hr>

�ȳ��ϼ���. <%= Request("UserID") %> ��.<br>
�����ID : <%= Request("UserID") %> <br> 
��й�ȣ : <%= Request("Password") %><br>
<%
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(Server.MapPath(".")+"\password.dat",1)

If Request("UserID") = Trim(objFile.ReadLine) AND Request("Password") = Trim(objFile.ReadLine) Then
	Session("id") = "Admin"
Else
	Session("id") = "Guest"
End IF

If Session("id") = "Admin" Then
	Response.Write("�����ڱ���...")
Else
	Response.Write("���� �մ����� �������Դϴ�.")	
End If
%>
<% End if %>
</body>    
</html>
