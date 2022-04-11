<%	
strTest = "안녕하세요.<p>"

Sub PYJ()
	'지역변수 : 해당 프로시저 안에서만 사용
	strTest = "반갑습니다.<p>"

	Response.Write strTest

End Sub

Response.Write strTest

PYJ()
%>
