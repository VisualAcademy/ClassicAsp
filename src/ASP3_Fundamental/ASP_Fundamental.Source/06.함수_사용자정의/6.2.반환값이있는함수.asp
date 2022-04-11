<script language="javascript">
//함수 선언부
function sum(a, b)
{
     strResult = parseInt(a) + parseInt(b);
     return strResult;
}
//함수 호출부
strResult=sum(3, 5);
document.write(strResult + "<br>");     
</script>
<%
'함수 선언부
Function Sum(a, b)
     strResult = CInt(a) + CInt(b)
     Sum = strResult
End Function
'함수 호출부
strResult=Sum(3, 5)
Response.Write(strResult & "<br>")
%>