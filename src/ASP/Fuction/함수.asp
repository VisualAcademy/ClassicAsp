<%
' 함수 선언부
Sub Hi()
    Response.Write("안녕<br />")
End Sub
Sub Hello(s)
    Response.Write(s & "<br />")
End Sub
Function Power(i)
    Power = (i * i)
End Function
%>
<%
' 함수 호출부
'[1] 매개변수도 없고 반환값도 없는 함수
Call Hi()
Hi() ' 권장
Hi
'[2] 매개변수가 있는 함수(메서드)
Hello("안녕")
Hello("방가")
Hello "또봐" ' 권장하지 않음
'[3] 반환값(Return Value)이 있는 함수
Response.Write( Power(2) & "<br />" ) ' 2의 제곱 = 4
%>
<%
Sub ValueReference(ByVal a, ByRef b)
    a = 5
    b = 10        
End Sub
%>
<%
Dim a
Dim b
a = 10
b = 20
Call ValueReference(a, b)
Response.Write("a : " & a & "<br />") ' 10
Response.Write("b : " & b & "<br />") ' 20->10
%>