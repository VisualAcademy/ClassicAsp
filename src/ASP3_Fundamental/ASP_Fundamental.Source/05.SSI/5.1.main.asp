<%
'�������̵� ��Ŭ���...SSI ���...
%>
<!-- #include file="./hi.asp" -->
<hr>
<!-- #include file="./hi.inc" -->
<hr>
<!-- #include file="./hi.txt" -->
<%
Call Hi()

Server.Execute "./hi.asp"	'hi.asp�� ����(���α׷�����)
Server.Transfer "./hi.inc"	'hi.inc�� ����(���α׷�����)
%>
