<%
'��ȯ���� �ִ� �Լ� ����
Function Hap(intA, intB)
	Result = CInt(intA) + CInt(intB)
	Hap = Result   '��ȯ���� �ѱ� ���� �Լ��̸��� ��ȯ���� �ش�.
End Function
%>


<%
'��ȯ���� �ִ� �Լ� ȣ��
Result=Hap(3,5)

Response.Write "�μ��� ���� " & Result & "�Դϴ�."
%>


