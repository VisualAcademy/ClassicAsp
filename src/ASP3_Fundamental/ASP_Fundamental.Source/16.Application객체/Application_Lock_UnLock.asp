<%
	Application.Lock   '���� ������ ���´�.

	'Check��� ���ø����̼� ������ �����ϰ�, �� ������ ����������
	'1������Ų��.
	Application("Check") = Application("Check") + 1

	Application.UnLock   '������ ���� Ǯ���ش�.

	Session("Check") = Session("Check") + 1
%>
�������� ������ ���� ����(Application����) : <%=Application("Check")%><br>
�������� ������ ���� ����(Session����) : <%=Session("Check")%>






