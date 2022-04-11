<%
Dim intSum
Dim i
intSum = 0
For i = 1 To 100 Step 1
	If i MOD 2 = 0 Then
		intSum = intSum + i
	End If
Next
Response.Write("1부터 100까지의 합 : " & intSum)
%>

<script language="javascript">
var intSum, i;
intSum = 0;
for(i = 1; i<=100; ++i)
{
	if(i % 2 == 0)
	{
		intSum += i;//intSum = intSum + i;
	}
}
document.write("1부터 100까지 짝수의 합 : " + intSum);
</script>
