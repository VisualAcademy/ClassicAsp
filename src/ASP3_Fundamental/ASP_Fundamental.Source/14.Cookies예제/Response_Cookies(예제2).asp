<html>
<body bgColor=#ffffff>
<% if request("userid") = "" then %>
	<p> �� ������ �Է� </p>
	<hr>
		<form action="Response_Cookies(����2).asp" method="post"> 
	   ����� ID : <input type="text" name="userid" size=20  value="<%=Request.Cookies("login")("id")%>"> <br>
	   ��й�ȣ : <input type="password" name="password" size=20 value="<%=Request.Cookies("login")("password")%>"><br>
	   ��й�ȣ ���� : <input type="checkbox" name="savepassword" value="yes"
	   <% If Request.Cookies("login")("password") <> "" Then 
				Response.Write "Checked"
			End If
	   %>
	   ><br>
	   <input type="submit" value="������"><input type="reset" value="�ʱ�ȭ">
	</form>
<% else %>
	<p> �Է��� ��� </p>
	<hr>
	�ȳ��ϼ���. <%= Request("UserID") %> ��.<br> �����ID : <%= Request("UserID") %> <br>  ��й�ȣ : <%= Request("Password") %>

	<% 
		Response.Cookies("login").Expires = Date() + 365   '��Ű�Ҹ�ð� ����(1��)
		'Response.Cookies("login").Expires = #09-24-2001#   '��Ű�Ҹ�ð�(��¥����)
		Response.Cookies("login").Domain = "xpplus.net"   '��뵵���� ����
		Response.Cookies("login").Path = "/"   '��뵵���γ� Ư�����丮 ����

		'��Ű ����
	    Response.Cookies("login")( "id" ) = Request("UserID")
		If Request.Form("savepassword") = "" Then
			Response.Cookies("login")("password") = ""
		Else
			Response.Cookies("login")( "password" ) = Request("password")
		End If
	%>
<% end if %>
</body>    
</html>
