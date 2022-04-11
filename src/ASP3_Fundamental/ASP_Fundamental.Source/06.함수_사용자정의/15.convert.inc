<% '요일값을 반환시켜주는 함수
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