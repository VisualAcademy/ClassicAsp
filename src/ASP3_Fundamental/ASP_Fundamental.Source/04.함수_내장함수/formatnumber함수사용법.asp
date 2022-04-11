<html><head>
<title>formatnumber함수사용법</title>
</head><body bgcolor="#FFFFFF"><html>
<%
	'My ASP program that formats every permutation of a number 
	response.write "original number=5000000<P>"
	response.write "FormatCurrency(5000000,0)=" & formatcurrency(5000000) & "<P>"
	response.write "FormatCurrency(5000000,2)=" & formatcurrency(5000000,2) & "<P><HR>"

	response.write "original number=50<P>"
	response.write "FormatCurrency(50)=" & formatcurrency(50) & "<P>"
	response.write "FormatCurrency(50,0,tristatetrue)=" & formatcurrency(50,tristatetrue) & "<P><HR>"

	response.write "original number=-50<P>"
	response.write "FormatCurrency(-50)="
	response.write formatcurrency(-50) & "<P>"
	response.write "FormatCurrency(-50,0,tristateusedefault,tristateusedefault,tristatefalse)=" 
	response.write formatcurrency(50,tristateusedefault,tristateusedefault,tristatefalse) & "<P><HR>"
%>
</body></html>