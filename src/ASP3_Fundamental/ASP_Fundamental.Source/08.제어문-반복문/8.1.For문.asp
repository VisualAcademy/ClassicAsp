<%
'ASP로 1부터 100까지의 짝수의 합 : For문
Dim i
Dim sum
i = 0
Sum = 0

For i = 0 To 100 'Step 1
     If i MOD 2 = 0 Then
          Sum = Sum + i
     End If
Next

Response.Write("1부터 100까지의 짝수의 합 : " & Sum & "<br>")
%>

<script language="javascript">
//자바스크립트로 1부터 100까지의 짝수의 합 : for문 사용
var i, sum;
i = 0;
sum = 0;

for(i = 0; i <=100 ; i++)
{
     if(i % 2 == 0)
     {
          sum += i;  //sum = sum + i;
     }
}

document.write("1부터 100까지의 짝수의 합 : " + sum + "<br>");
</script>