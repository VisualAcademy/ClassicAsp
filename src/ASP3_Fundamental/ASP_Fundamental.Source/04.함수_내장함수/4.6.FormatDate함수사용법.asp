<html>
<head>
<title>formatdate함수사용법</title>
</head>
<body bgcolor="#FFFFFF">
<%
Response.Write "<hr>"
For counter=0 To 4
	currentdate=now()
	Response.Write "today is..." & "<br>"
	Response.Write currentdate & "<P>"
	Select Case counter
		Case 0
			whichformat="vbgeneraldate"
		Case 1
			whichformat="vblongdate"
		Case 2
			whichformat="vbshortdate"
		Case 3
			whichformat="vblongtime"
		Case 4
			whichformat="vbshorttime"
	End Select
	Response.Write "FormatDate(now()," & whichformat & ")="
	Response.Write Formatdatetime(currentdate,counter) & "<P><HR>"
Next
%>
</body>
</html>