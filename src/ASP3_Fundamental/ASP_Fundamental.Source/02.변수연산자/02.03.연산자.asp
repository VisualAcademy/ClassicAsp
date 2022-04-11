<%
	'[1] 변수의 명시적 선언
	Option Explicit

	'연산자
	'CDbl(), CInt(), CStr() 함수사용

	'[2] 변수 선언
	Dim intA, dblB, numC

	'[3] 변수값 대입 및 처리
	intA = 10                     'integer형
	dblB = "10.5"            'string 형
	numC = intA + dblB   'ASP는 변수의 유형에 맞추어 (+)연산을 한다.

	'[4] 결과 출력
	Response.Write dblB & TypeName(dblB) & "<br>"

	numC = CInt(intA) + CInt(dblB)
	Response.Write numC & TypeName(dblB) & "<br>"

	numC = CDbl(intA) + CDbl(dblB)
	Response.Write numC & TypeName(numC) & "<br>"

	numC = CStr(intA) + CStr(dblB)
	Response.Write numC & TypeName(numC) & "<br>"
%>
