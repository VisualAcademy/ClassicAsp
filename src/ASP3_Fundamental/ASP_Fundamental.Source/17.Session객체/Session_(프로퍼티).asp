<h4>Session Ÿ�Ӿƿ� </h4>
<hr>
    Session.SessionID(OS�����) : <%=Session.SessionID%> <br>
    Session.Timeout(���������ð�) : <%=Session.Timeout%> �� <br>

    Session.CodePage(���) : <%= Session.CodePage%><br>
    Session.LCID(����) : <%= Session.LCID %>

    <% Session.Timeout = 20 %>
    <br> ���� Ÿ�� �ƿ��� <%= Session.Timeout %> ������ �ٲ�
