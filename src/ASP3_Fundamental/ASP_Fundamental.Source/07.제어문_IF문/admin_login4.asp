<HTML>
<HEAD>
<TITLE>������ �α���</TITLE>
</HEAD>
<BODY BGCOLOR="#FFFFFF">
<div align="center">
<table border="1">
	<form name="login_form" method="post" action="./admin_login4_check.asp">
	<tr>
		<td align="center">
			������ �α���
		</td>
	</tr>

	<% If Session("id") = "admin" Then %>
	<tr>
		<td align="center">
			���� �����ڷ� ���� ���Դϴ�.
		</td>
	</tr>
	<% Else %>
	<tr>
		<td align="center">
			���� �մ����� ���� ���Դϴ�.
		</td>
	</tr>
	<% End If %>

	<tr>
		<td align="center">
			���̵�:<input type="text" name="id">
		</td>
	</tr>
	<tr>
		<td align="center">
			�н�����:<input type="password" name="pwd">
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="submit" value="������ �α����ϱ�">
		</td>
	</tr>
	</form>
</table>
</div>
</BODY>
</HTML>
