<%
' 1���� 100���� 3�� ��� �Ǵ� 4�� ����� ���� ���ϴ� ���α׷�
Option Explicit ' ������ �������� �ʰ� ����ϸ� �����߻�
Dim sum: sum = 0
Dim i
i = 1
Do
    If i Mod 3 = 0 Or i Mod 4 = 0 Then
        sum = sum + i
    End If
    i = i + 1
Loop While (i <= 100)
Response.Write("����� �� : " & sum & "<br />")
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
document.write("����� �� : " + sum + "<br />");
</script>