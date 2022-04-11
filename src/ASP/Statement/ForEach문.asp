<%
Dim a(2)

a(0) = "홍길동"
a(1) = "백두산"
a(2) = "한라산"

' 0번째부터 2번째 인덱스까지 반복
Dim i
For i = 0 To 2 
    Response.Write(a(i) & "<br />")
Next

' a라는 배열에 데이터가 있는만큼 반복 : 3번 반복
Dim s
For Each s In a
    Response.Write(s & "<br />")
Next
%>