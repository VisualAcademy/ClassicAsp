<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_admin.asp
' Program Description : 폼태그를 이용한 관리자 로그인
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<HTML>
	<HEAD>
		<TITLE>관리자 로그인 페이지</TITLE>
		<script language="javascript">
function write_form_check()
{
	//아이디 입력 체크
   if(window.document.write_form.id.value == "")
   {
      window.alert("비밀번호를 넣어주세요.");   //메시지 박스 출력.
	  window.document.write_form.id.focus();   //해당 입력양식에 포커스.
	  window.document.write_form.id.select();   //해당 입력양식값을 선택.
	  return false;   //현재 함수를 호출한 곳으로 false값 반환.
   }
   //비밀번호 입력 체크
   if(window.document.write_form.passwd.value == "")
   {
      window.alert("비밀번호를 넣어주세요.");   //메시지 박스 출력.
	  window.document.write_form.passwd.focus();   //해당 입력양식에 포커스.
	  window.document.write_form.passwd.select();   //해당 입력양식값을 선택.
	  return false;   //현재 함수를 호출한 곳으로 false값 반환.
   }
   window.document.write_form.submit();   //체크사항 만족시 전송.
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
									▣관리자 비밀번호▣
								</P>
							</td>
						</tr>
						<tr>
							<td>
								<P align="center">
									아이디(초기값:admin)&nbsp;
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
									비밀번호(초기값:1111)&nbsp;
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
									<input type="button" value="관리자 로그인 하기" onclick="write_form_check();" ID="Button1" NAME="Button1">&nbsp;&nbsp;
									<input type="reset" value="뒤&nbsp;&nbsp;로" onClick="history.go(-1);" ID="Reset1" NAME="Reset1">
								</P>
							</td>
						</tr>
				</form>
				</TBODY>
			</table>
		</DIV>
	</BODY>
</HTML>
