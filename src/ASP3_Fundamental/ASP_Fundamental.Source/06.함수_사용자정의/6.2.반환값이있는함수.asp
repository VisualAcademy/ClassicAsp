<script language="javascript">
//�Լ� �����
function sum(a, b)
{
     strResult = parseInt(a) + parseInt(b);
     return strResult;
}
//�Լ� ȣ���
strResult=sum(3, 5);
document.write(strResult + "<br>");     
</script>
<%
'�Լ� �����
Function Sum(a, b)
     strResult = CInt(a) + CInt(b)
     Sum = strResult
End Function
'�Լ� ȣ���
strResult=Sum(3, 5)
Response.Write(strResult & "<br>")
%>