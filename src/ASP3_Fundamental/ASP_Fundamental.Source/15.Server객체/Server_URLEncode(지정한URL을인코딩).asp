<%
StrURL="http://127.0.0.1/test.asp?key=<' >"
Response.Write "<p>���ڵ� �� : " & StrURL

'URL�� ���ڵ��մϴ�.
StrURLEn = Server.URLEncode(StrURL)
Response.Write "<p>���ڵ� �� : " & StrURLEn
%>
