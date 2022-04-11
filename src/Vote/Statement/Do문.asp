<script>
var i = 1;
do
{
    document.write(i + "번째<br />");
    i++;
} while(i <= 10);
</script>
<%
Dim i
i = 1
Do
    Response.Write(i & "번째<br />")
    i = i + 1
Loop While i <= 10 ' 10보다 작을 동안
i = 1
Do
    Response.Write(i & "번째<br />")
    i = i + 1
Loop Until i > 10 ' 10보다 클 때까지
i = 1
Do Until i > 10 ' 10보다 클 때까지
    Response.Write(i & "번째<br />")
    i = i + 1
Loop 
i = 1
Do While i <= 10 ' 10보다 작을 동안
    Response.Write(i & "번째<br />")
    i = i + 1
Loop 
%>