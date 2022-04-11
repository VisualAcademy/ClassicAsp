<h4> Execute 메서드 </h4>
<p>
<h5>Server.Execute("Server_MapPath(지정된실제경로반환).asp") 를 부르기 전... </h5>
<hr>
<% 
    Server.Execute "./Server_MapPath(지정된실제경로반환).asp"    '제어권이 다시 돌아옴.
	'Server.Transfer "./Server_MapPath(지정된실제경로반환).asp"   '제어권을 넘김.
%>

<hr>
<h5>Server.Execute() 실행이 끝난 후</h5>

<!-- #include file="./Server_MapPath(지정된실제경로반환).asp" -->
