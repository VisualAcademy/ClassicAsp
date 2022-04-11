<%
Sub EvenNumber(n)
    Dim sum
    Dim i
    sum = 0
    For i = 1 To n
        If i Mod 2 = 0 Then
            sum = sum + i
        End If
    Next
    Response.Write("합계 : " & sum & "<br />")
End Sub
%>
<%
EvenNumber(100) ' 1부터 100까지 짝수의 합
EvenNumber(50) ' 1부터 50까지 짝수의 합
EvenNumber(10) ' 1부터 10까지 짝수의 합
%>