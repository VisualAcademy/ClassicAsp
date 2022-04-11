<%
'[1] 지역 변수
Dim a
a = 10
'[2] 전역 변수
Application("a") = 10
Application("b") = 20
Application("c") = 30
%>
<%= "a = " & a %><br />
<%= "Application(""a"") = " & Application("a") %><br />

<%
'Application.Contents.Remove("Now")
For Each n In Application.Contents
    Response.Write(Application(n) & "<br />")
Next
%>