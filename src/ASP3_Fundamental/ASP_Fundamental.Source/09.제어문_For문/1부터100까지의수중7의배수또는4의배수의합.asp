<%
Dim Sum, i

Sum = 0

For i = 1 To 100
	If	(i MOD 7 = 0) OR (i MOD 4 = 0) Then
		Sum = Sum + i
	End If
Next

Response.Write("1부터 100까지 7의 배수 또는 4의 배수의 합 : " & Sum)

%>

<script language="javascript">
var sum, i;

sum=0;

for(i=1;i<=100;i++)
{
	if(i % 7 == 0 || i % 4 == 0)
	{
		sum += i;
	}
}

document.write("1부터 100까지 7의 배수 또는 4의 배수의 합 : " + sum);
</script>


