<br><br><br>

<TABLE border="1" width="500" align="center">

<form action="./BoardDeleteProcess.asp" method="post">

<input type="hidden" name="Num" value="<%=Request("Num")%>">

<TR>
	<TD colspan="2"><%=Request("Num")%>번 글을 삭제하시겠습니까?</TD>
</TR>
<TR>
	<TD><%=Request("Num")%>번 글 비밀번호</TD>
	<TD><input type="password" name="Password"></TD>
</TR>
<TR>
	<TD colspan="2" align="center">
		<input type="submit" value="삭제">
	</TD>
</TR>

</form>

</TABLE>