<%
Option Explicit
Dim objCon: Dim objRs
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASP;Initial Catalog=ASP;Data Source=.\SQLEXPRESS")
'//Open(SQL문, 커넥션, 커서, 락, 옵션)

Set objRs = Server.CreateObject("ADODB.RecordSet")
objRs.Open "Select * From Memos Order By Num Desc", objCon
If objRs.BOF Or objRs.EOF Then
	Response.Write("데이터가 없습니다.")
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
<center><h3>간단한 한줄 메모장</h3>
<!-- #include file="./MemoWrite.asp" -->
<TABLE border=1 align="center">
<tr>
<td>번호</td><td>이름</td><td>메모</td><td>작성일</td>
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


