<% If Request("userid") = "" Then %>

������ ������ �Է�<br>
<hr>
<form action="" method="post"> 
   ����� ID : <input type="text" name="UserID" size=20><br>
   ��й�ȣ : <input type="password" name="Password" size=20><p>
   <input type="submit" value="������">
   <input type="reset" value="�ʱ�ȭ">
</form>

<% Else %>

���۵� ������ ���<br>
<hr>
�ȳ��ϼ���. <%= Request.Form("UserID") %> ��.<br>
�����ID : <%= Request.QueryString("UserID") %> <br> 
��й�ȣ : <%= Request("Password") %>

<% End If %>