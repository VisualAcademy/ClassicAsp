<script>
var i = 1;
do
{
    document.write(i + "��°<br />");
    i++;
} while(i <= 10);
</script>
<%
Dim i
i = 1
Do
    Response.Write(i & "��°<br />")
    i = i + 1
Loop While i <= 10 ' 10���� ���� ����
i = 1
Do
    Response.Write(i & "��°<br />")
    i = i + 1
Loop Until i > 10 ' 10���� Ŭ ������
i = 1
Do Until i > 10 ' 10���� Ŭ ������
    Response.Write(i & "��°<br />")
    i = i + 1
Loop 
i = 1
Do While i <= 10 ' 10���� ���� ����
    Response.Write(i & "��°<br />")
    i = i + 1
Loop 
%>