<%
Dim a(2)

a(0) = "ȫ�浿"
a(1) = "��λ�"
a(2) = "�Ѷ��"

' 0��°���� 2��° �ε������� �ݺ�
Dim i
For i = 0 To 2 
    Response.Write(a(i) & "<br />")
Next

' a��� �迭�� �����Ͱ� �ִ¸�ŭ �ݺ� : 3�� �ݺ�
Dim s
For Each s In a
    Response.Write(s & "<br />")
Next
%>