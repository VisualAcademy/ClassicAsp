<script language="javascript">
//�Լ� �����
function sum(a, b)
{
     strResult = parseInt(a) + parseInt(b);
               document.write("�ڹٽ�ũ��Ʈ ��� : " + strResult + "<br>");     
}
//�Լ� ȣ���
sum(3, 5);
</script>
<%
'�Լ� �����
Function Sum(a, b)
     strResult = CInt(a) + CInt(b)
                Response.Write("ASP ��� : " & strResult & "<br>")
End Function
'�Լ� ȣ���
Call Sum(3, 5)
%>