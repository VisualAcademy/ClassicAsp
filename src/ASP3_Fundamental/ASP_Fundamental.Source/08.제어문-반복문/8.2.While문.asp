<%
'ASP로 1부터 100까지의 짝수의 합 : While문
Dim i
Dim sum
i = 0
Sum = 0
i = 0                                  '초기값
While i <= 100                         '조건식
     If i MOD 2 = 0 Then
          Sum = Sum + i
     End If
     i = i + 1                         '증감식
WEnd

Response.Write("1부터 100까지의 짝수의 합 : " & Sum & "<br>")
%>


<script language="javascript">
//자바스크립트로 1부터 100까지의 짝수의 합 : while문 사용
var i, sum;
i = 0;
sum = 0;
i = 0;
while(i <=100)
{
     if(i % 2 == 0)
     {
          sum += i;  //sum = sum + i;
     }
     i++;
}

document.write("1부터 100까지의 짝수의 합 : " + sum + "<br>");
</script>