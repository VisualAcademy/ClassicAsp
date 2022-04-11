<HTML>
<HEAD>
<TITLE>관리자 로그인</TITLE>
</HEAD>
<BODY BGCOLOR="#FFFFFF">
<div align="center">
<table border="1">
	<form name="login_form" method="post" action="./admin_login4_check.asp">
	<tr>
		<td align="center">
			관리자 로그인
		</td>
	</tr>

	<% If Session("id") = "admin" Then %>
	<tr>
		<td align="center">
			현재 관리자로 접속 중입니다.
		</td>
	</tr>
	<% Else %>
	<tr>
		<td align="center">
			현재 손님으로 접속 중입니다.
		</td>
	</tr>
	<% End If %>

	<tr>
		<td align="center">
			아이디:<input type="text" name="id">
		</td>
	</tr>
	<tr>
		<td align="center">
			패스워드:<input type="password" name="pwd">
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="submit" value="관리자 로그인하기">
		</td>
	</tr>
	</form>
</table>
</div>
</BODY>
</HTML>
