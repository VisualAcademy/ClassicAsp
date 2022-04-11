<%
'반환값이 있는 함수 선언
Function Hap(intA, intB)
	Result = CInt(intA) + CInt(intB)
	Hap = Result   '반환값을 넘길 때는 함수이름에 반환값을 준다.
End Function
%>


<%
'반환값이 있는 함수 호출
Result=Hap(3,5)

Response.Write "두수의 합은 " & Result & "입니다."
%>


