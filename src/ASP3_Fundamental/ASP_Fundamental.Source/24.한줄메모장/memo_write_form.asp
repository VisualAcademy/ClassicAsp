<%
'--------------------------------------------------------------------
' Title : MyMemo Education Version
' Program Name : memo_write_form.asp
' Program Description : �޸�� �����ϴ� �Է���
' Include Files : None
' Last Update : 2001.08.25. 23:20
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<HTML>
	<HEAD>
		<TITLE>MY�޸��� �Է� ������</TITLE>
		<SCRIPT language="javascript">
function write_form_check()
{
   //�̸��� �Է� üũ
   if(window.document.write_form.name.value == "")
   {
      window.alert("�̸��� �־��ּ���.");   //�޽��� �ڽ� ���.
	  window.document.write_form.name.focus();   //�ش� �Է¾�Ŀ� ��Ŀ��.
	  window.document.write_form.name.select();   //�ش� �Է¾�İ��� ����.
	  return false;   //���� �Լ��� ȣ���� ������ false�� ��ȯ.
   }
   //�̸� ���̰� üũ
   if(window.document.write_form.name.value.length < 2 || window.document.write_form.name.value.length > 5)
   {
      window.alert("�̸��� 2�ڸ��̻� 5�ڸ� ���Ϸ� �־��ּ���.");   //�޽��� �ڽ� ���.
	  window.document.write_form.name.focus();   //�ش� �Է¾�Ŀ� ��Ŀ��.
	  window.document.write_form.name.select();   //�ش� �Է¾�İ��� ����.
	  return false;   //���� �Լ��� ȣ���� ������ false�� ��ȯ.
   }
   //�̸��ϰ� �Է� üũ
   if(window.document.write_form.email.value == "")
   {
      window.alert("�̸����� �־��ּ���.");   //�޽��� �ڽ� ���.
	  window.document.write_form.email.focus();   //�ش� �Է¾�Ŀ� ��Ŀ��.
	  window.document.write_form.email.select();   //�ش� �Է¾�İ��� ����.
	  return false;   //���� �Լ��� ȣ���� ������ false�� ��ȯ.
   }
   //�̸��� ���̰� �Է� üũ
   if(window.document.write_form.email.value.length < 10 || window.document.write_form.email.value.length > 30)
   {
      window.alert("�̸����� ��Ȯ�ϰ� �־��ּ���.");   //�޽��� �ڽ� ���.
	  window.document.write_form.email.focus();   //�ش� �Է¾�Ŀ� ��Ŀ��.
	  window.document.write_form.email.select();   //�ش� �Է¾�İ��� ����.
	  return false;   //���� �Լ��� ȣ���� ������ false�� ��ȯ.
   }
   //�޸� �Է� üũ
   if(window.document.write_form.title.value == "")
   {
      window.alert("�޸� �����ּ���.");   //�޽��� �ڽ� ���.
	  window.document.write_form.title.focus();   //�ش� �Է¾�Ŀ� ��Ŀ��.
	  window.document.write_form.title.select();   //�ش� �Է¾�İ��� ����.
	  return false;   //���� �Լ��� ȣ���� ������ false�� ��ȯ.
   }
   //�޸� ���̰� �Է� üũ
   if(window.document.write_form.title.value.length < 1 || window.document.write_form.title.value.length > 50)
   {
      window.alert("�޸��� ���̸� 50�����Ϸ� �־��ּ���.");   //�޽��� �ڽ� ���.
	  window.document.write_form.title.focus();   //�ش� �Է¾�Ŀ� ��Ŀ��.
	  window.document.write_form.title.select();   //�ش� �Է¾�İ��� ����.
	  return false;   //���� �Լ��� ȣ���� ������ false�� ��ȯ.
   }
   window.document.write_form.submit();   //üũ���� ������ ����.
}
		</SCRIPT>
	</HEAD>
	<BODY bgcolor="#FFFFFF">
		<DIV align="center">
			<TABLE border="1" width="100%">
				<FORM name="write_form" method="post" action="./memo_write_process.asp">
					<TR>
						<TD colspan="2" align="center" align="center">
							�ø޸� ������
						</TD>
					</TR>
					<TR>
						<TD align="center">
							��&nbsp;&nbsp;&nbsp;��
						</TD>
						<TD align="center">
							<INPUT type="text" name="name" size="30">
						</TD>
					</TR>
					<TR>
						<TD align="center">
							�̸���
						</TD>
						<TD align="center">
							<INPUT type="text" name="email" size="30">
						</TD>
					</TR>
					<TR>
						<TD align="center">
							��&nbsp;&nbsp;&nbsp;��
						</TD>
						<TD align="center">
							<INPUT type="text" name="title" size="30">
						</TD>
					</TR>
					<TR>
						<TD colspan="2" align="center">
							<INPUT type="button" value="�۾���" onclick="write_form_check();">&nbsp;&nbsp; 
							<INPUT type="reset" value="��&nbsp;&nbsp;��">
						</TD>
					</TR>
				</FORM>
			</TABLE>
		</DIV>
	</BODY>
</HTML>
