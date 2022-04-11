<%
StrURL="http://127.0.0.1/test.asp?key=<' >"
Response.Write "<p>인코딩 전 : " & StrURL

'URL을 인코딩합니다.
StrURLEn = Server.URLEncode(StrURL)
Response.Write "<p>인코딩 후 : " & StrURLEn
%>
