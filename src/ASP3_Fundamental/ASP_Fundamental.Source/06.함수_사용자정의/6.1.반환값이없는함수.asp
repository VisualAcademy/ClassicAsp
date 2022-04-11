<script language="javascript">
//함수 선언부
function sum(a, b)
{
     strResult = parseInt(a) + parseInt(b);
               document.write("자바스크립트 사용 : " + strResult + "<br>");     
}
//함수 호출부
sum(3, 5);
</script>
<%
'함수 선언부
Function Sum(a, b)
     strResult = CInt(a) + CInt(b)
                Response.Write("ASP 사용 : " & strResult & "<br>")
End Function
'함수 호출부
Call Sum(3, 5)
%>