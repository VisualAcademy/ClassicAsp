<%
Set objNewMessage = Server.CreateObject("CDO.Message")

objNewMessage.From = "redplus@redplus.net"     '//������ ���
objNewMessage.To = "redplus@redplus.net"     '//�޴� ���
objNewMessage.Subject = "2003���� ������ �׽�Ʈ"     '//����
objNewMessage.HtmlBody = "<b>HTML�� ��ȯ����</b>"     '//����

objNewMessage.Send()                                        '//����

Set objNewMessage = Nothing
%>
<center>���� ���� �Ϸ�</center> 