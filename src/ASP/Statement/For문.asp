<script>
var i;
for (i = 1; i <= 10; i++)
{
    document.write(i + "��°<br />");
}
</script>
<%
Dim i
For i = 1 To 10 Step 1
    Response.Write(i & "��°<br />")
Next
%>