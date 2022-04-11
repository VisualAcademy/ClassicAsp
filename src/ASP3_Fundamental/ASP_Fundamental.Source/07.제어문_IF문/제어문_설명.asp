<%
'While문
Dim Sum, Count
Sum = 0
Count =1

While Count <= 100
	If Count MOD 2 = 0 Then
		Sum = Sum + Count
	End If
	Count = Count + 1
WEnd

Response.Write("1~100까지 짝수의 합 : " & Sum)
%>
<script language="javascript">
//while문
var sum, count;
sum = 0;
count = 1;

while(count<=100)
{
	if(count % 2 == 0)
	{
		sum += count;	
	}
	count += 1;
}

window.alert("1~100까지 짝수의 합 : " + sum);
</script>


