<%
Option Explicit
Dim a: Dim b

a = "10"
b = 10.5

Response.Write( a + b & "<br>" ) '?
Response.Write( CInt(a) + b & "<br>") 'CInt() : 무늬만 숫자인것을 숫자로
Response.Write( a + CStr(b) & "<br>") 'CStr() : 문자열로 변환 1010.5

Response.Write( CDbl(a) + b )  'CDbl() : 실수형으로 변환
%>