<% if request("userid") = "" then %>

<p> �� ������ �Է� </p>
<hr>

<form action="" method="post"> 
   ����� ID : <input type="text" name="userid" size=20><br>
   ��й�ȣ : <input type="password" name="password" size=20><p>
   <input type="submit" value="������">
   <input type="reset" value="�ʱ�ȭ">
</form>

<% else %>

<p> �Է��� ��� </p>
<hr>

�ȳ��ϼ���. <%= Request("UserID") %> ��.<br>
�����ID : <%= Request("UserID") %> <br> 
��й�ȣ : <%= Request("Password") %>

<% end if %>
