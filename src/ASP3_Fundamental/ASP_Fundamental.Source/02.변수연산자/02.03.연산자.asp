<%
	'[1] ������ ����� ����
	Option Explicit

	'������
	'CDbl(), CInt(), CStr() �Լ����

	'[2] ���� ����
	Dim intA, dblB, numC

	'[3] ������ ���� �� ó��
	intA = 10                     'integer��
	dblB = "10.5"            'string ��
	numC = intA + dblB   'ASP�� ������ ������ ���߾� (+)������ �Ѵ�.

	'[4] ��� ���
	Response.Write dblB & TypeName(dblB) & "<br>"

	numC = CInt(intA) + CInt(dblB)
	Response.Write numC & TypeName(dblB) & "<br>"

	numC = CDbl(intA) + CDbl(dblB)
	Response.Write numC & TypeName(numC) & "<br>"

	numC = CStr(intA) + CStr(dblB)
	Response.Write numC & TypeName(numC) & "<br>"
%>
