<%
'--------------------------------------------------------------------
' MyMemo Education Version
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<HTML>
	<HEAD>
		<TITLE>관리지비밀번호입력</TITLE>
	</HEAD>
	<BODY bgcolor="#FFFFFF">
		<DIV align="center">
			<TABLE border="1" width="100%">
				<FORM method="post" action="./memo_delete_process.asp">
					<TR>
						<TD align="center">
							관리자비밀번호(초기값:1234)
						</TD>
					<TR>
						<TD align="center">
							<INPUT type="password" name="passwd">
						</TD>
					</TR>
					<TR>
						<TD align="center">
							지울글의 번호
						</TD>
					</TR>
					<TR>
						<TD align="center">
							<INPUT type="num" name="num">
						</TD>
					</TR>
					<TR>
						<TD align="center">
							<INPUT type="submit" value="지우기">&nbsp;&nbsp;&nbsp;<INPUT type="button" value="뒤&nbsp;&nbsp;로" onclick="history.go(-1);">
						</TD>
					</TR>
				</FORM>
			</TABLE>
		</DIV>
	</BODY>
</HTML>
