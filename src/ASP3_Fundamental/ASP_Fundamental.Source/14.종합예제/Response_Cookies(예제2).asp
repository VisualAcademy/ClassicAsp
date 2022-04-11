<html>
<body bgColor=#ffffff>
<% if request("userid") = "" then %>
	<p> 폼 데이터 입력 </p>
	<hr>
		<form action="Response_Cookies(예제2).asp" method="post"> 
	   사용자 ID : <input type="text" name="userid" size=20  value="<%=Request.Cookies("login")("id")%>"> <br>
	   비밀번호 : <input type="password" name="password" size=20 value="<%=Request.Cookies("login")("password")%>"><br>
	   비밀번호 저장 : <input type="checkbox" name="savepassword" value="yes"
	   <% If Request.Cookies("login")("password") <> "" Then 
				Response.Write "Checked"
			End If
	   %>
	   ><br>
	   <input type="submit" value="보내기"><input type="reset" value="초기화">
	</form>
<% else %>
	<p> 입력한 결과 </p>
	<hr>
	안녕하세요. <%= Request("UserID") %> 님.<br> 사용자ID : <%= Request("UserID") %> <br>  비밀번호 : <%= Request("Password") %>

	<% 
		Response.Cookies("login").Expires = Date() + 365   '쿠키소멸시간 설정(1년)
		'Response.Cookies("login").Expires = #09-24-2001#   '쿠키소멸시간(날짜지정)
		Response.Cookies("login").Domain = "xpplus.net"   '사용도메인 지정
		Response.Cookies("login").Path = "/"   '사용도메인내 특정디렉토리 지정

		'쿠키 저장
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
