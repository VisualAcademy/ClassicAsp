<%
	Option Explicit
	Dim intSince
	'내장함수 : 날짜 관련 함수
	'Now(), Date(), Time()
	Response.Write Now()   '2001-08-21 오후 6:16:33 → 사용빈도 높음.
	Response.Write("<br>")
	Response.Write Now   '2001-08-21 오후 6:16:33 → 사용빈도 높음.
	Response.Write("<br>")
	Response.Write Date()   '2001-08-21
	Response.Write("<br>")
	Response.Write Time()   '오후 6:17:53
	Response.Write("<br>")

	'날짜 추출 함수
	'Year(), Month(), Day(), WeekDay(), Hour(), Minute(), Second()
	Response.Write(Year(Now()))   '년
	Response.Write("<br>")
	Response.Write(Month(Now()))   '월
	Response.Write("<br>")
	Response.Write(Day(Now()))   '일
	Response.Write("<br>")
	Response.Write(WeekDay(Now()))   '요일(1:일요일 ~ 7:토요일)
	Response.Write("<br>")
	Response.Write(Hour(Now()))   '시간
	Response.Write("<br>")
	Response.Write(Minute(Now()))   '분
	Response.Write("<br>")
	Response.Write(Second(Now()))   '초
	Response.Write("<br>")

	'날짜에 대한 숫자값을 반환(DateSerial()).
	intSince = Date() - DateSerial(2002,01,01)
	Response.Write("올해가 벌써 " & intSince &" 일이 지났습니다.")

   '크리스마스 앞으로 며칠이 남았는지요???
	intSince = DateSerial(2002,12,25) - Date()
	Response.Write("<br>")
	Response.Write("크리스마스 앞으로 " & intSince &"일이 남았습니다.")

%>
	FormatDateTime(Date,1)의 결과 : <%=FormatDateTime(Date,1)%><br>
	FormatDateTime(Date,2)의 결과 : <%=FormatDateTime(Date,2)%><br>
	FormatDateTime(Date,3)의 결과 : <%=FormatDateTime(Date,3)%><br>
	FormatDateTime(Date,4)의 결과 : <%=FormatDateTime(Date,4)%><p>
