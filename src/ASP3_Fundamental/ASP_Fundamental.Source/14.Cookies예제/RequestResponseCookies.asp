<% If Request("UserID") = "" Then %>
	<form name="" action="" method="post"> 
		����� ID : <input type="text" name="UserID" size=20  
			value="<%=Request.Cookies("id")%>"><br>
		��й�ȣ : <input type="password" name="Password" size=20 
			value="<%=Request.Cookies("Password")%>"><br>
		��й�ȣ ���� : <input type="checkbox" name="SavedPWD" 
			value="yes"
		<% 
		If Request.Cookies("Password") <> "" Then 
			Response.Write "Checked"
		End If
		%>><br>
		<input type="submit" value="������"><input type="reset" value="�ʱ�ȭ">
	</form>
<% 
Else 
	Response.Cookies( "UserID" ) = Request("UserID")
	If Request.Form("SavedPWD") = "" Then
		Response.Cookies("Password") = ""
	Else
		Response.Cookies( "Password" ) = Request("Password")
	End If
	Response.Write("��Ű�� ����Ǿ����ϴ�.")
End If 
%>
