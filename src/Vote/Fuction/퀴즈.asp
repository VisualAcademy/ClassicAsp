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
    Response.Write("�հ� : " & sum & "<br />")
End Sub
%>
<%
EvenNumber(100) ' 1���� 100���� ¦���� ��
EvenNumber(50) ' 1���� 50���� ¦���� ��
EvenNumber(10) ' 1���� 10���� ¦���� ��
%>