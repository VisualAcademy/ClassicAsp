<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_delete.asp
' Program Description : ���±׸� �̿��� ������� �н����� ����
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
Option Explicit

'���� ����
Dim Num

'�Ѱ����� ���ڵ� �ѹ�(Num)���� ������ ����
Num = Request.QueryString("Num")   '���� : QueryString �÷������� �޾ƾ� ��.
%>
<HTML>
	<HEAD>
		<TITLE>�Խ��� ���� ������</TITLE>
		<script language="javascript">
function write_form_check()
{
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
				<form name="write_form" method="post" action="./board_delete_process.asp" ID="Form1">
				<input type="hidden" name="Num" value="<%=Num%>">
					<TBODY>
						<tr>
							<td colspan="2" align="middle">
								<P align="center">
									��������� ��й�ȣ �Է¢�
								</P>
							</td>
						</tr>
						<tr>
							<td>
								<P align="center">
									��й�ȣ&nbsp;
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
									<input type="button" value="�����" onclick="write_form_check();" ID="Button1" NAME="Button1">&nbsp;&nbsp;
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
