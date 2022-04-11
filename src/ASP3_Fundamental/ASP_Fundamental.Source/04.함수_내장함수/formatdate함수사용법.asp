<html><head>
<title>formatdate함수사용법</title>
</head><body bgcolor="#FFFFFF"><html>
<%
response.write "<hr>"
for counter=0 to 4
	currentdate=now()
	response.write "today is..." & "<br>"
	response.write currentdate & "<P>"
	select case counter
	case 0
		whichformat="vbgeneraldate"
	case 1
		whichformat="vblongdate"
	case 2
		whichformat="vbshortdate"
	case 3
		whichformat="vblongtime"
	case 4
		whichformat="vbshorttime"
	end select
	response.write "FormatDate(now()," & whichformat & ")="
	response.write Formatdatetime(currentdate,counter) & "<P><HR>"
next%>
</body></html>