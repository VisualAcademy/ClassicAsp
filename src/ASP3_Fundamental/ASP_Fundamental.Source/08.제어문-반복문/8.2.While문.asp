<%
'ASP�� 1���� 100������ ¦���� �� : While��
Dim i
Dim sum
i = 0
Sum = 0
i = 0                                  '�ʱⰪ
While i <= 100                         '���ǽ�
     If i MOD 2 = 0 Then
          Sum = Sum + i
     End If
     i = i + 1                         '������
WEnd

Response.Write("1���� 100������ ¦���� �� : " & Sum & "<br>")
%>


<script language="javascript">
//�ڹٽ�ũ��Ʈ�� 1���� 100������ ¦���� �� : while�� ���
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

document.write("1���� 100������ ¦���� �� : " + sum + "<br>");
</script>