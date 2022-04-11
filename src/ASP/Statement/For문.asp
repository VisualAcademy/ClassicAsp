<script>
var i;
for (i = 1; i <= 10; i++)
{
    document.write(i + "번째<br />");
}
</script>
<%
Dim i
For i = 1 To 10 Step 1
    Response.Write(i & "번째<br />")
Next
%>