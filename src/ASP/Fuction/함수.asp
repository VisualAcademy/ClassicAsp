<%
' �Լ� �����
Sub Hi()
    Response.Write("�ȳ�<br />")
End Sub
Sub Hello(s)
    Response.Write(s & "<br />")
End Sub
Function Power(i)
    Power = (i * i)
End Function
%>
<%
' �Լ� ȣ���
'[1] �Ű������� ���� ��ȯ���� ���� �Լ�
Call Hi()
Hi() ' ����
Hi
'[2] �Ű������� �ִ� �Լ�(�޼���)
Hello("�ȳ�")
Hello("�氡")
Hello "�Ǻ�" ' �������� ����
'[3] ��ȯ��(Return Value)�� �ִ� �Լ�
Response.Write( Power(2) & "<br />" ) ' 2�� ���� = 4
%>
<%
Sub ValueReference(ByVal a, ByRef b)
    a = 5
    b = 10        
End Sub
%>
<%
Dim a
Dim b
a = 10
b = 20
Call ValueReference(a, b)
Response.Write("a : " & a & "<br />") ' 10
Response.Write("b : " & b & "<br />") ' 20->10
%>