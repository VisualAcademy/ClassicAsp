<%
	'���� ����
	Dim id, pwd

	'���۵� �н����� ���� Request��ü�� Form�÷������� �޾Ƽ� pwd������ ����
	id = LCase(Request.Form("id"))   '���̵�� ��ҹ��� ���� ������...
	pwd = Request.Form("pwd")

	'���࿡ ���۵� �н����尪�� 1234(��������)�̸� ������ �޽��� ����ϰ�,
	'�׷��� ������, �ڹٽ�ũ��Ʈ�� ����ؼ� ���� ������ �̵��Ѵ�.
	If id="admin" AND pwd = "1234" Then
		Response.Write("<center>�������Դϴ�.</center>")
	Else
%>
	<script language="JavaScript">
		window.alert("�����ڰ� �ƴϱ���\n�ٽ� �α����ϼ���\n���� ��ŷ��???");
		history.go(-1);   //history.back();
	</script>
<%
	End If
%>
		