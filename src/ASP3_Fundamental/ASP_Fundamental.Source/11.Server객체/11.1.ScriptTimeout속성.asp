<%
    Response.Write("���� ��ũ��Ʈ Ÿ�Ӿƿ� �ð� : " & _
		Server.ScriptTimeout & "��<p>")

    ' ��ũ��Ʈ Ÿ�Ӿƿ� �ð��� 3������ �ٲ�
    Server.ScriptTimeout = 180

    Response.Write("�ٲ� ��ũ��Ʈ Ÿ�Ӿƿ� �ð� : " & _
		Server.ScriptTimeout & "��<p>")
%>