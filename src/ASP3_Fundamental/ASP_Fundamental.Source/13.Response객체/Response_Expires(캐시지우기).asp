<%
	Response.Expires = -1	'현재 페이지를 매번 새로 읽어온다.
	'Response.Expires = 60	'1시간뒤에 캐시소멸

	Response.ExpiresAbsolute = #12/25/2002 00:00:00#   '2002년 12월 25일에 캐시 소멸.
%>

이곳의 내용을 웹브라우저는 캐시에 저장한다.<br>
...
