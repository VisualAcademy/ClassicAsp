<br><br><br>

<TABLE border="1" width="500" align="center">

<form action="./BoardDeleteProcess.asp" method="post">

<input type="hidden" name="Num" value="<%=Request("Num")%>">

<TR>
	<TD colspan="2"><%=Request("Num")%>�� ���� �����Ͻðڽ��ϱ�?</TD>
</TR>
<TR>
	<TD><%=Request("Num")%>�� �� ��й�ȣ</TD>
	<TD><input type="password" name="Password"></TD>
</TR>
<TR>
	<TD colspan="2" align="center">
		<input type="submit" value="����">
	</TD>
</TR>

</form>

</TABLE>