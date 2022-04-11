<%
' 시간대마다 서로 다른 메시지 출력하는 프로그램
Dim msg
Select Case Hour(Now())
    Case 16
        msg = "나른한 오후"
    Case 17
        msg = "오후 5시"
    Case Else
        msg = "다른 시간대"
End Select
Response.Write(msg)
%>
<script>
var msg;
var today = new Date();
switch (today.getHours())
{
    case 17: msg = "오후 5시대 메시지"; break;
    default: msg = "다른 시간대"; break; 
}
document.write(msg);
</script>