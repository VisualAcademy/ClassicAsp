<%
Option Explicit
Dim objCon: Dim objRs
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASP;Initial Catalog=ASP;Data Source=.\SQLEXPRESS")
'//Open(SQL��, Ŀ�ؼ�, Ŀ��, ��, �ɼ�)

Set objRs = Server.CreateObject("ADODB.RecordSet")
objRs.Open "Select * From Memos Order By Num Desc", objCon
If objRs.BOF Or objRs.EOF Then
	Response.Write("�����Ͱ� �����ϴ�.")
Else
	Call ShowRecordSet(objRs)
End If
objRs.Close()
Set objRs = Nothing

objCon.Close()
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(ByRef objRs)
%>
<center><h3>������ ���� �޸���</h3>
<!-- #include file="./MemoWrite.asp" -->
<TABLE border=1 align="center">
<tr>
<td>��ȣ</td><td>�̸�</td><td>�޸�</td><td>�ۼ���</td>
</tr>
<%
	While Not objRs.EOF
%>
<TR onmouseover="this.style.backgroundColor='red';"
onmouseout="this.style.backgroundColor='';">
	<TD>
	<a href="./MemoDelete.asp?Num=<%=objRs("Num")%>">
	<%=objRs.Fields("Num")%>
	</a>
	</TD>
	<TD><%=objRs("Name")%></TD>
	<TD><%=objRs.Fields(3)%></TD>
	<TD><%=objRs(4)%></TD>
</TR>
<%
		objRs.MoveNext()
	WEnd
%>
</TABLE>
<%
End Sub
%>


