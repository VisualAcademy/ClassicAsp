<% If Request("UserID") = "" Then %>
	<form name="" action="" method="post"> 
		사용자 ID : <input type="text" name="UserID" size=20  
			value="<%=Request.Cookies("id")%>"><br>
		비밀번호 : <input type="password" name="Password" size=20 
			value="<%=Request.Cookies("Password")%>"><br>
		비밀번호 저장 : <input type="checkbox" name="SavedPWD" 
			value="yes"
		<% 
		If Request.Cookies("Password") <> "" Then 
			Response.Write "Checked"
		End If
		%>><br>
		<input type="submit" value="보내기"><input type="reset" value="초기화">
	</form>
<% 
Else 
	Response.Cookies( "UserID" ) = Request("UserID")
	If Request.Form("SavedPWD") = "" Then
		Response.Cookies("Password") = ""
	Else
		Response.Cookies( "Password" ) = Request("Password")
	End If
	Response.Write("쿠키가 저장되었습니다.")
End If 
%>
