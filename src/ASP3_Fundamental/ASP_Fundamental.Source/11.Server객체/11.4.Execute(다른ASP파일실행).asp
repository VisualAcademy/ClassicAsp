<!-- #include file="./hi.asp" -->
<% 
    Server.Execute "./hi.asp"    '������� �ٽ� ���ƿ�.
	Server.Transfer "./hi.asp"   '������� �ѱ�.
    Server.Execute "./hi.asp"    '���� �ȵ�
%>