<%
' �ð��븶�� ���� �ٸ� �޽��� ����ϴ� ���α׷�
Dim msg
Select Case Hour(Now())
    Case 16
        msg = "������ ����"
    Case 17
        msg = "���� 5��"
    Case Else
        msg = "�ٸ� �ð���"
End Select
Response.Write(msg)
%>
<script>
var msg;
var today = new Date();
switch (today.getHours())
{
    case 17: msg = "���� 5�ô� �޽���"; break;
    default: msg = "�ٸ� �ð���"; break; 
}
document.write(msg);
</script>