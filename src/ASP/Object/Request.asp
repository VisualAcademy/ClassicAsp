<%
' ���̵� �ޱ�
Dim UserID
Dim Password
UserID = Request.Form("UserID")
Password = Request("Password")
%>
�Ѱ��� �� ���̵�� <%= UserID %>�̰�,
��ȣ�� <%= Password %><br />
IP�ּҴ� <%= Request.ServerVariables("REMOTE_ADDR") %><br />
IP�ּҴ� <%= Request.ServerVariables("REMOTE_HOST") %><br />
