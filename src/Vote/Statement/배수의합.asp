<%
' 1부터 100까지 3의 배수 또는 4의 배수의 합을 구하는 프로그램
Option Explicit ' 변수를 선언하지 않고 사용하면 에러발생
Dim sum: sum = 0
Dim i
i = 1
Do
    If i Mod 3 = 0 Or i Mod 4 = 0 Then
        sum = sum + i
    End If
    i = i + 1
Loop While (i <= 100)
Response.Write("배수의 합 : " & sum & "<br />")
%>
<script>
var sum = 0;
var i;
for (i = 1; i <= 100; i++)
{
    if (i % 3 == 0 || i % 4 == 0)
    {
        sum += i;
    }
}
document.write("배수의 합 : " + sum + "<br />");
</script>