<%
	'[1] 변수 명시적 선언 : 현재 페이지에서 변수를 하나라도 선언하지 않고 사용하면 에러를 발생시킨다.
	'주석을 제외한 첫줄에 와야한다.
	Option Explicit

	'[2] 변수 선언 : Dim이라는 키워드를 사용해서 선언한다.
	'헝가리안 표기법 또는 파스칼표기법을 사용해서 변수명을 정한다.
	Dim strHi   '문자형 변수...(String)
	Dim intSu   '정수형 변수...(Integer)
	Dim objDB   '객체형 변수...(Object)
	Dim dateBirthday   '날짜형 변수...
	Dim strSQL, strDB   '변수를 여러개 동시에 선언하려면 ,(쉼표)를 사용해서 나열한다.

	'[3] 변수 값 대입
	strHi= "안녕하세요"   'ASP에서 문자열 대입은 "(큰따옴표)로 묶는다.
	intSu = 100
	dateBirthday = #08-21-2001#   '날짜값을 대입하려면 날짜 문자열 앞뒤로 #(Sharp) 기호를 붙인다.

	'[4] 변수 값 출력
	Response.Write("strHi : " & strHi & "<br>" )
	Response.Write "intSu+intSu : " & intSu+intSu & "<br>"
	Response.Write "dateBirthday : " & dateBirthday & "<br>"
	
	'[5] 변수 유형 값 출력(ASP내장함수인 TypeName() 사용 : 해당 변수의 유형을 표시한다.)
	Response.Write(strHi & ": &nbsp;" & TypeName(strHi) & "<br>")
	Response.Write(intSu & ":&nbsp;" & TypeName(intSu+intSu) & "<br>")
	Response.Write(dateBirthday & ":&nbsp;" & TypeName(dateBirthday))
%>