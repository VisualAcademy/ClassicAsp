<%
Option Explicit
Dim a: Dim b

a = "10"
b = 10.5

Response.Write( a + b & "<br>" ) '?
Response.Write( CInt(a) + b & "<br>") 'CInt() : ���̸� �����ΰ��� ���ڷ�
Response.Write( a + CStr(b) & "<br>") 'CStr() : ���ڿ��� ��ȯ 1010.5

Response.Write( CDbl(a) + b )  'CDbl() : �Ǽ������� ��ȯ
%>