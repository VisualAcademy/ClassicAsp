<%
'서버사이드 인클루드...SSI 사용...
%>
<!-- #include file="./hi.asp" -->
<hr>
<!-- #include file="./hi.inc" -->
<hr>
<!-- #include file="./hi.txt" -->
<%
Call Hi()

Server.Execute "./hi.asp"	'hi.asp를 실행(프로그램지속)
Server.Transfer "./hi.inc"	'hi.inc를 실행(프로그램종료)
%>
