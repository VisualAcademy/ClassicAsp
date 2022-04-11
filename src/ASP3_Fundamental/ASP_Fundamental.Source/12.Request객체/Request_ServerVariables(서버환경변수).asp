<html>
<head>
<title>당신의 현재 사용중인 시스템의 정보입니다!</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
	A:link		{font: 9pt GulimChe,굴림체; color: navy; text-decoration: none}
	A:active	{font: 9pt GulimChe,굴림체; color: red;  text-decoration: none}
	A:visited	{font: 9pt GulimChe,굴림체; color: navy; text-decoration: none}
	A:hover		{							color: #0000ff; text-decoration: none}
		td		{font: 9pt GulimChe,굴림체; color: black;}
		.bold	{font: bold 13pt GulimChe,굴림체; color: white;}
-->
</style>
</head>
<body>
	<p>
	<table border=1>
		<%	For Each key in Request.ServerVariables %>
		<tr>
			<td>	<%= key %>	</td>
			<td>
				<% If Request.ServerVariables(key) = "" then
						Response.Write "&nbsp;"
					else
						Response.Write Request.ServerVariables(key)
					end if
					
					Response.Write	"</td>"
				%>
		</tr>
		<% next %>
	</table>
	<p>
</body>
</html>
