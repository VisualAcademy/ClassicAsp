<%
Response.Buffer = True ' ���۸� ���(��� �����ְڴ�..)

Response.Write("[1] �ȳ�<br />")

Response.Flush() ' ��������� ���� ���� ���

Response.Write("[2] ��¾ȵ�<br />")

Response.Clear() ' ���� ���� ������ ����

Response.Write("[3] ���<br />")

Response.End() ' ��������� ���� ��� �� ����

Response.Write("[4] ��µ��� ����<br />")
%>