<%
' 산술연산자 : +, -, *, /, Mod
'[1] 변수 선언
Dim a
Dim b
Dim c
'[2] 변수 초기화
a = 10
b = 5
c = b Mod a ' b Mod a -> b / a : 몫 : 0, 나머지 : 5
'[3] 변수 출력
Response.Write(c & "<br />") ' 5
'[4] 변수의 유형 출력
Response.Write(TypeName(c) & "<br />") ' Integer    
%>
<script>
var a = 10; var b = 5; var c = b % a;
window.document.write(c + "<br />");
document.write(typeof(c) + "<br />"); // number
</script>