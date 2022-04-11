<html>
<head>
<title>시스템 정보</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
	td {font-size:8pt;font-family:verdana;}
</style>
</head>
<body>
	<table border=1>
		<%For Each key in Request.ServerVariables%>
		<tr>
			<td><%=key%></td>
			<td>
				<%
				If Request.ServerVariables(key) = "" Then
					Response.Write "&nbsp;"
				Else
					Response.Write Request.ServerVariables(key)
				End If
					Response.Write	"</td>"
				%>
		</tr>
		<%Next%>
	</table>
</body>
</html>