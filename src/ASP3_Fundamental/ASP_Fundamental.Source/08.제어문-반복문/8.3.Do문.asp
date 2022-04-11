<%
     'Do~Loop문으로 1~100까지의 수 중 짝수의 합을 구하는 프로그램
     Dim Sum, Count

     Sum = 0
     Count = 1
     Do Until Count > 100   '조건식이 거짓일 동안...(참일 때까지)
          If Count Mod 2 = 0 Then
               Sum = Sum + Count
          End If
          Count = Count + 1   '카운터 변수 증가
     Loop

     Response.Write("<center>1부터 100까지의 짝수의 합은 " & Sum & " 입니다</center>")

     Sum = 0
     Count = 1
     Do While Count <= 100   '조건식이 참일동안...
          If Count Mod 2 = 0 Then
               Sum = Sum + Count
          End If
          Count = Count + 1   '카운터 변수 증가
     Loop

     Response.Write("<center>1부터 100까지의 짝수의 합은 " & Sum & " 입니다</center>")

     Sum = 0
     Count = 1
     Do    '최소한 한번은 실행한다.
          If Count Mod 2 = 0 Then
               Sum = Sum + Count
          End If
          Count = Count + 1   '카운터 변수 증가
     Loop Until Count > 100   '조건식이 거짓일 동안...(참일 때까지)

     Response.Write("<center>1부터 100까지의 짝수의 합은 " & Sum & " 입니다</center>")

     Sum = 0
     Count = 1
     Do    '최소한 한번은 실행한다.
          If Count Mod 2 = 0 Then
               Sum = Sum + Count
          End If
          Count = Count + 1   '카운터 변수 증가
     Loop While Count <= 100   '조건식이 참일 동안...(거짓일 때까지)

     Response.Write("<center>1부터 100까지의 짝수의 합은 " & Sum & " 입니다</center>")
%>