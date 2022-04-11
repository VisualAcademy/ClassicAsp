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
						alert("받는 사람의 메일 주소를 입력하세요.");
						f.tb_receiver.focus();
						return false;
					}
					else {
						var receivers = receiver.split(";");
						for (var i=0;i<receivers.length;i++) {
							if (receivers[i] != "") {
								if (!mail_regex1.test(receivers[i])) {
									alert("입력하신 메일 주소 '" + receivers[i] + "'는 잘못된 메일 주소입니다.");
									f.tb_receiver.focus();
									return false;
								}

								if (mail_regex2.test(receivers[i])) {
									alert("입력하신 메일 주소 '" + receivers[i] + "'는 사용할 수 없는 메일 주소입니다.");
									f.tb_receiver.focus();
									return false;
								}
							}
						}
					}

					if (sender == "") {
						alert("보내는 사람의 메일 주소를 입력하세요.");
						f.tb_sender.focus();
						return false;
					}
					else {
						if (!mail_regex1.test(sender)) {
							alert("보내는 사람의 메일 주소가 잘못되었습니다.");
							f.tb_sender.focus();
							return false;
						}
					}

					if (attach != "") {
						var ext = attach.substring(attach.lastIndexOf(".") + 1).toLowerCase();
						
						if ((ext == "asp") || (ext == "jsp") ||
							(ext == "php") || (ext == "php3") || (ext == "phtml") ||
							(ext == "exe") || (ext == "bat") || (ext == "com")) {
							alert("선택하신 파일 형식은 첨부할 수 없습니다.");
							return false;
						}
					}

					if (subject == "") {
						if (!confirm("제목이 입력되지 않았습니다.\n제목없이 보내시겠습니까?")) {
							f.tb_subject.focus();
							return false;
						}
					}

					if (content == "") {
						alert("내용을 입력하세요.");
						f.tb_content.focus();
						return false;
					}

					return true;
				}
				catch (ex) {
					if (confirm("실행 중에 다음과 같은 오류가 발생하였습니다.\n\n\"" + ex.message + "\"\n\n페이지를 새로 고침하시겠습니까?"))
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
								<b>아래의 항목들을 입력해 주세요.</b><br><br>
								여러 명에게 메일을 보내시려면 받는 사람의 메일 주소를<br>세미 콜론으로 구분하세요.<br>
								ex) id1@domain.com; id2@domain.com
							</td>
						</tr>
					</table>
					<table border="0" cellpadding="2" cellspacing="1" width="100%">
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>받는 사람:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="text" name="tb_receiver" size="40" class="input">
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>보내는 사람:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="text" name="tb_sender" size="20" class="input" value="<% = Request.Cookies("mail_remember_me") %>">
								<input type="checkbox" name="cb_remember"<% If Len(Request.Cookies("mail_remember_me")) > 0 Then Response.Write " checked" %>>내 메일 주소 기억
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								<b>파일 첨부:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="file" name="tb_attach" size="25" class="input">
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>메일 형식:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="radio" name="radio_bodyformat" value="1" checked>일반 텍스트
								<input type="radio" name="radio_bodyformat" value="0">HTML
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>제목:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<input type="text" name="tb_subject" size="40" class="input">
							</td>
						</tr>
						<tr>
							<td width="106" bgcolor="#DDDDD5">
								&nbsp;<b>내용:</b>
							</td>
							<td width="250" bgcolor="#EEEEE5">
								<textarea name="tb_content" cols="40" rows="10" class="input"></textarea>
							</td>
						</tr>
						<tr>
							<td colspan="2" align="center">
								<input type="submit" name="btn_send" value="메일 전송" class="btn" onfocus="blur()">
								<input type="reset" name="btn_cancel" value="다시 작성" class="btn" onfocus="blur()">
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