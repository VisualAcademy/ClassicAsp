<html>
<head>
<title>http://www.xpplus.net/</title>
<body bgcolor="#FFFFFF">
</head>
Do-Loop�� ���� : 1������ ����ޱ����� ���� ���.<br>
<%
counter = 1
thismonth = month(now())

Do while counter < thismonth + 1
	response.write "month number " & counter & " "
	response.write "______________________________" & "<BR><br>"

	If counter >13 then
		exit do
	end if
	counter = counter+1
Loop

%>
<hr>
</body>
</html>
