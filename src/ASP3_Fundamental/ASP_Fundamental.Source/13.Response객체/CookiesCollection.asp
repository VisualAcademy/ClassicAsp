<%
'Response��ü�� Cookies�÷����� ����� ��Ű ����
    ' �⺻Ű�� ����� ��Ű ����
    Response.Cookies( "ID" ) = "redplus1"
    Response.Cookies( "PWD" ) = "1234"
    ' ����Ű�� ����� ��Ű ����
    Response.Cookies("User")("ID") = "redplus2"
    Response.Cookies("User")("PWD") = "5678"
'Request��ü�� Cookies�÷����� ����� ��Ű �о����
	Response.Write("ID : " & Request.Cookies("ID") & "<br>")
	Response.Write("PWD : " & Request.Cookies("PWD") & "<br>")
	Response.Write("User.ID : " & Request.Cookies("User")("ID") & "<br>")
	Response.Write("User.PWD : " & Request.Cookies("User")("PWD") & "<br>")
%>