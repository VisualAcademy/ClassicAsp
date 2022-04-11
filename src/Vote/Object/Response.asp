<%
Response.Buffer = True ' 버퍼링 사용(끊어서 보여주겠다..)

Response.Write("[1] 안녕<br />")

Response.Flush() ' 현재까지의 버퍼 내용 출력

Response.Write("[2] 출력안됨<br />")

Response.Clear() ' 현재 버퍼 내용을 삭제

Response.Write("[3] 출력<br />")

Response.End() ' 현재까지의 내용 출력 후 종료

Response.Write("[4] 출력되지 않음<br />")
%>