<html>
<body bgcolor=White>

<% if request("mode") = "out" then %>
<p> �Է��� ��� </p>
<hr>

�ȳ��ϼ���. <%= Request("UserID") %> ��.<br>
�����ID : <%= Request("UserID") %> <br> 
��й�ȣ : <%= Request("Password") %>

<% else %>

<p> �� ������ �Է� </p>
<hr>

<form action="Request_Hidden_Form.asp" method="Get"> 

   ����� ID : <input type="text" name="userid" size=20><br>
   ��й�ȣ : <input type="password" name="password" size=20><p>

   <input type="submit" value="������">
   <input type="reset" value="�ʱ�ȭ">

   <input type="hidden" name="mode" value="out">
</form>

<% end if %>
</body>    
</html>

