<%@ Language=VBScript %>
<html>
	<head>
		<title>E-Mail Deliverer version 0.1</title>
		<style type="text/css">
			body, table, td, input, span, div, textarea, option {
				font-family:tahoma;
				font-size:8pt;
			}

			.input {
				border:solid 1 #696969;
			}

			.btn {
				border:solid 1 #DDDDD5;
				background-color:#DDDDDD5;
				font-weight:bold;
				color:#2D2D2D;
			}
		</style>
		<script language="javascript">
			function validate_form() {
				try {
					var f = document.forms[0];
					var mail_regex1 = /[\w|-]+@[\w|-|]+[.]+[\w|-|.]+/;
					var mail_regex2 = /hanmail.net+|daum.net+|lycos.co.kr+/;

					var receiver = f.tb_receiver.value.replace(" ", "");
					var sender = f.tb_sender.value;
					var attach = f.tb_attach.value;
					var subject = f.tb_subject.value;
					var content = f.tb_content.value;

					if (receiver == "") {
						alert("�޴� ����� ���� �ּҸ� �Է��ϼ���.");
						f.tb_receiver.focus();
						return false;
					}
					else {
						var receivers = receiver.split(";");
						for (var i=0;i<receivers.length;i++) {
							if (receivers[i] != "") {
								if (!mail_regex1.test(receivers[i])) {
									alert("�Է��Ͻ� ���� �ּ� '" + receivers[i] + "'�� �߸��� ���� �ּ��Դϴ�.");
									f.tb_receiver.focus();
									return false;
								}

								if (mail_regex2.test(receivers[i])) {
									alert("�Է��Ͻ� ���� �ּ� '" + receivers[i] + "'�� ����� �� ���� ���� �ּ��Դϴ�.");
									f.tb_receiver.focus();
									return false;
								}
							}
						}
					}

					if (sender == "") {
						alert("������ ����� ���� �ּҸ� �Է��ϼ���.");
						f.tb_sender.focus();
						return false;
					}
					else {
						if (!mail_regex1.test(sender)) {
							alert("������ ����� ���� �ּҰ� �߸��Ǿ����ϴ�.");
							f.tb_sender.focus();
							return false;
						}
					}

					if (attach != "") {
						var ext = attach.substring(attach.lastIndexOf(".") + 1).toLowerCase();
						
						if ((ext == "asp") || (ext == "jsp") ||
							(ext == "php") || (ext == "php3") || (ext == "phtml") ||
							(ext == "exe") || (ext == "bat") || (ext == "com")) {
							alert("�����Ͻ� ���� ������ ÷���� �� �����ϴ�.");
							return false;
						}
					}

					if (subject == "") {
						if (!confirm("������ �Էµ��� �ʾҽ��ϴ�.\n������� �����ðڽ��ϱ�?")) {
							f.tb_subject.focus();
							return false;
						}
					}

					if (content == "") {
						alert("������ �Է��ϼ���.");
						f.tb_content.focus();
						return false;
					}

					return true;
				}
				catch (ex) {
					if (confirm("���� �߿� ������ ���� ������ �߻��Ͽ����ϴ�.\n\n\"" + ex.message + "\"\n\n�������� ���� ��ħ�Ͻðڽ��ϱ�?"))
						location.reload();
					else
						return false;
				}
			}
		</script>
	</head>
	<body>
		<form name="form1" method="post" action="SendMail.asp" onsubmit="return validate_form()"
		enctype="multipart/form-data">
		<table border="0" cellpadding="0" cellspacing="0" width="559" align="center">
			<tr>
				<td width="559" colspan="3"><img src="imgs/box_top.gif" border="0"></td>
			</tr>
			<tr>
				<td width="185"><img src="imgs/box_left.gif" border="0"></td>
				<td width="356">
					<table border="0" cellpadding="1" cellspacing="1" width="100%">
						<tr>
							<td width="100%">
								<b>�Ʒ��� �׸���� �Է��� �ּ���.</b><br><br>
								���� ���� ������ �����÷��� �޴� ����� ���� �ּҸ�<br>���� �ݷ����� �����ϼ���.<br>
								ex) id1@domain.com; id2@domain.com
							</td>
						</tr>
					</table>
					<table border="0" cellpadding="2" cellspacing="1" width="100%">
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>�޴� ���:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="text" name="tb_receiver" size="40" class="input">
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>������ ���:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="text" name="tb_sender" size="20" class="input" value="<% = Request.Cookies("mail_remember_me") %>">
								<input type="checkbox" name="cb_remember"<% If Len(Request.Cookies("mail_remember_me")) > 0 Then Response.Write " checked" %>>�� ���� �ּ� ���
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								<b>���� ÷��:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="file" name="tb_attach" size="25" class="input">
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>���� ����:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="radio" name="radio_bodyformat" value="1" checked>�Ϲ� �ؽ�Ʈ
								<input type="radio" name="radio_bodyformat" value="0">HTML
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>����:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="text" name="tb_subject" size="40" class="input">
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>����:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<textarea name="tb_content" cols="40" rows="10" class="input"></textarea>
							</td>
						</tr>
						<tr>
							<td colspan="2" align="center">
								<input type="submit" name="btn_send" value="���� ����" class="btn" onfocus="blur()">
								<input type="reset" name="btn_cancel" value="�ٽ� �ۼ�" class="btn" onfocus="blur()">
							</td>
						</tr>
					</table>
				</td>
				<td width="18"><img src="imgs/box_right.gif" border="0"></td>
			</tr>
			<tr>
				<td width="559" colspan="3"><img src="imgs/box_bottom.gif" border="0"></td>
			</tr>
		</table>
		</form>
	</body>
</html>