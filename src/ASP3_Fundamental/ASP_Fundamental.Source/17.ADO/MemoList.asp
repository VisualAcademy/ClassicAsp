<%
Option Explicit
Dim objCon: Dim objRs
Set objCon = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.RecordSet")
objCon.Open("Provider=SQLOLEDB.1;Password=memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")
'//Open(SQL��, Ŀ�ؼ�, Ŀ��, ��, �ɼ�)
objRs.Open "Select * From Memo Order By Num Desc", objCon
If objRs.BOF Or objRs.EOF Then
	Response.Write("�����Ͱ� �����ϴ�.")
Else
	Call ShowRecordSet(objRs)
End If
objCon.Close()
Set objCon = Nothing
%>

<%
Sub ShowRecordSet(objRs)
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


