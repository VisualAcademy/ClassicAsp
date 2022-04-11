<!-- #include file="./hi.asp" -->
<% 
    Server.Execute "./hi.asp"    '제어권이 다시 돌아옴.
	Server.Transfer "./hi.asp"   '제어권을 넘김.
    Server.Execute "./hi.asp"    '실행 안됨
%>