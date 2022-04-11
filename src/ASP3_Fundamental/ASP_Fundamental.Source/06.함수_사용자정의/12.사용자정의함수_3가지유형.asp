<%@ Language=VBScript %>
<%
'-------------------------------------
' 사용자정의 함수 예제 - 작성자 : 박용준
'-------------------------------------
%>

<% '함수 기본 모양 : 서브 프로시저
SUB Hi() 

	 Response.Write("서브프로시저 출력")
	 
END SUB 
%>

<% '매개변수값을 받아서 처리하는 함수
SUB PrintHTML2(printsize, printcolor) 
%>

	<font size = "<%=printsize%>" color="<%=printcolor%>">
	여기는 사용자 정의 함수 출력부분 - 매개변수값에 따른 출력
	</font>

<% 
END SUB 
%>

<% '요일값을 반환시켜주는 함수 : 값을 반환시킬 때 함수이름에 값을 넣는다. : 함수
Function ConvertToWeekDay(str_week) 

     select case str_week
          case 1 : ConvertToWeekDay = "일요일"
          case 2 : ConvertToWeekDay = "월요일"
          case 3 : ConvertToWeekDay = "화요일"
          case 4 : ConvertToWeekDay = "수요일"
          case 5 : ConvertToWeekDay = "목요일"
          case 6 : ConvertToWeekDay = "금요일"
          case 7 : ConvertToWeekDay = "토요일"
     end select
     
End Function
%>

사용자 정의 함수 예제
<HR>
<%
'[1] 매개변수(파라미터) 및 리턴값이 없는 함수 호출
CALL Hi() 
'Hi()
'Hi
%>
<HR>
<%
'[2] 매개변수가 있는 함수 호출
str_size = 4
str_color = "red"
CALL PrintHTML2(str_size,str_color) 
%>
<HR>
<% 
'[3] 매개변수 및 리턴값이 있는 함수 호출
str_week = WeekDay(Now)
Response.Write(ConvertToWeekDay(str_week)) 
%>
