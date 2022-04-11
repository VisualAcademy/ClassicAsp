<%
'Response객체의 Cookies컬렉션을 사용한 쿠키 저장
    ' 기본키를 사용한 쿠키 저장
    Response.Cookies( "ID" ) = "redplus1"
    Response.Cookies( "PWD" ) = "1234"
    ' 서브키를 사용한 쿠키 저장
    Response.Cookies("User")("ID") = "redplus2"
    Response.Cookies("User")("PWD") = "5678"
'Request객체의 Cookies컬렉션을 사용한 쿠키 읽어오기
	Response.Write("ID : " & Request.Cookies("ID") & "<br>")
	Response.Write("PWD : " & Request.Cookies("PWD") & "<br>")
	Response.Write("User.ID : " & Request.Cookies("User")("ID") & "<br>")
	Response.Write("User.PWD : " & Request.Cookies("User")("PWD") & "<br>")
%>