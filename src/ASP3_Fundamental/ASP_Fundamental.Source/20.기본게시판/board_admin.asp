<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_admin.asp
' Program Description : ���±׸� �̿��� ������ �α���
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<HTML>
	<HEAD>
		<TITLE>������ �α��� ������</TITLE>
		<script language="javascript">
function write_form_check()
{
	//���̵� �Է� üũ
   if(window.document.write_form.id.value == "")
   {
      window.alert("��й�ȣ�� �־��ּ���.");   //�޽��� �ڽ� ���.
	  window.document.write_form.id.focus();   //�ش� �Է¾�Ŀ� ��Ŀ��.
	  window.document.write_form.id.select();   //�ش� �Է¾�İ��� ����.
	  return false;   //���� �Լ��� ȣ���� ������ false�� ��ȯ.
   }
   //��й�ȣ �Է� üũ
   if(window.document.write_form.passwd.value == "")
   {
      window.alert("��й�ȣ�� �־��ּ���.");   //�޽��� �ڽ� ���.
	  window.document.write_form.passwd.focus();   //�ش� �Է¾�Ŀ� ��Ŀ��.
	  window.document.write_form.passwd.select();   //�ش� �Է¾�İ��� ����.
	  return false;   //���� �Լ��� ȣ���� ������ false�� ��ȯ.
   }
   window.document.write_form.submit();   //üũ���� ������ ����.
}
		</script>
	</HEAD>
	<BODY BGCOLOR="#ffffff">
		<DIV align="center">
			<table border="1">
				<form name="write_form" method="post" action="./board_admin_process.asp" ID="Form1">
					<TBODY>
						<tr>
							<td colspan="2" align="middle">
								<P align="center">
									�ð����� ��й�ȣ��
								</P>
							</td>
						</tr>
						<tr>
							<td>
								<P align="center">
									���̵�(�ʱⰪ:admin)&nbsp;
								</P>
							</td>
							<td>
								<P align="center">
									<input type="text" name="id" size="30">
								</P>
							</td>
						</tr>
						<tr>
							<td>
								<P align="center">
									��й�ȣ(�ʱⰪ:1111)&nbsp;
								</P>
							</td>
							<td>
								<P align="center">
									<input type="password" name="passwd" size="30" ID="Password1">
								</P>
							</td>
						</tr>
						<tr>
							<td colspan="2" align="middle">
								<P align="center">
									<input type="button" value="������ �α��� �ϱ�" onclick="write_form_check();" ID="Button1" NAME="Button1">&nbsp;&nbsp;
									<input type="reset" value="��&nbsp;&nbsp;��" onClick="history.go(-1);" ID="Reset1" NAME="Reset1">
								</P>
							</td>
						</tr>
				</form>
				</TBODY>
			</table>
		</DIV>
	</BODY>
</HTML>
