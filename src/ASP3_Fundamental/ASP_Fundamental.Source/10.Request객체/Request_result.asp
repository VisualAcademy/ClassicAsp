<HTML>
<BODY>
<font face="����" size="2">
<P>&nbsp;</P>
<center>
<h2>Request�� ���� ����</h2>
<p>&nbsp;</P>

�̸��� <%=Request.Form("username")%> �Դϴ�.<p>
���̴� <%=Request.Form("age")%> �Դϴ�.<p>
��ȭ��ȣ�� <%=Request("tel")%> �Դϴ�.<p>

'Request
'.form()  -- �÷���...=(�迭) : Ű���� ����ؼ� ���� ��ȯ...(post������� ���۽�)
'.QueryString() : get������� ���۽�(<A>)
'Request��ü�� �ΰ��� �÷����� ��������ʰ� �ٷ� ������ ���� �� �ִ�.
- Request("�Ѱܿ��̸�")

</center></font>
</BODY>
</HTML>
