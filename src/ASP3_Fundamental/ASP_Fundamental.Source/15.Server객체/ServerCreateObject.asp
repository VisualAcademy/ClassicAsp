<%
    Dim objBrowser

    ' ��ü ����	
	'Set ��ü�̸� = Server.CreateObject("���赵")
    Set objBrowser = Server.CreateObject("MSWC.BrowserType") 

    ' ��ü ���� Ȯ��
    If IsObject( objBrowser ) Then

        ' ������ ��ü�� ������Ƽ�� �޼��� �̿�  %>
        ����ϰ� �ִ� �������� <%=objBrowser.Browser%> <%=objBrowser.Version%> �Դϴ�. <P> 
<%  End If %>

