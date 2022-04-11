<% '매개변수값을 받아서 처리하는 함수
SUB PrintHTML2(printsize,printcolor) 
%>

     <HR>
     <font size = "<%=printsize%>" color="<%=printcolor%>">
     여기는 사용자 정의 함수 출력부분 - 매개변수값에 따른 출력
     </font>
     <HR>

<% 
END SUB 
%>

<%
 str_size = 4
str_color = "red"
' CALL PrintHTML2(str_size,str_color)
PrintHTML2 str_size,str_color   '매개변수가 있는 경우는 괄호 생략.
%>