<%
'ASP�� 1���� 100������ ¦���� �� : For��
Dim i
Dim sum
i = 0
Sum = 0

For i = 0 To 100 'Step 1
     If i MOD 2 = 0 Then
          Sum = Sum + i
     End If
Next

Response.Write("1���� 100������ ¦���� �� : " & Sum & "<br>")
%>

<script language="javascript">
//�ڹٽ�ũ��Ʈ�� 1���� 100������ ¦���� �� : for�� ���
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

document.write("1���� 100������ ¦���� �� : " + sum + "<br>");
</script>