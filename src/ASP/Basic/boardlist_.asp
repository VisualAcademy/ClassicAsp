<%
'--------------------------------------------------
' Title : Basic ����
' Program Name : boardlist.asp
' Program Description : �� ����Ʈ ������
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
'[1] ���� ����
'Option Explicit
Dim objCon: Dim objRs: Dim strSql
Dim Page: Page = 1'//
If Request("Page") <> "" Then
	Page = Request("Page")'�Ѱ����� ������ ���� ����.
End If
Dim PageCount'//�� ������ �� ����
'[2] �ν��Ͻ� ����
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] ����
objCon.Open(Application("ConnectionString"))
'[4] ���ڵ�� ��ü�� �ν��Ͻ� ����
Set objRs = Server.CreateObject("ADODB.RecordSet")
'[5] ��ɾ� ����
strSql = "Select * From Basic Order By Num Desc"
objRs.PageSize = 5'//������ ������ ����
objRs.Open strSql, objCon, 1'//1 �Ǵ� 3���� ����
'[6] ���
If objRs.BOF Or objRs.EOF Then
	Response.Write("�����Ͱ� �����ϴ�.")
Else
	PageCount = objRs.PageCount'//�� ������ �� ����
	objRs.AbsolutePage = Page'//
	Call ShowRecordSet(objRs)
End If
'[7] Close()
objRs.Close(): objCon.Close()
'[8] Nothing
Set objRs = Nothing: Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>

	<center><h3>�⺻ �Խ��� ����Ʈ</h3></center>

	<table border=1 width="100%">
	<tr>
		<TD>��ȣ</TD>
		<TD>����</TD>
		<TD>�ۼ���</TD>
		<TD>�ۼ���</TD>
		<TD>��ȸ��</TD>
	</tr>
<%
	Dim i: i = 1'//������ ������ ī��Ʈ. �ʱ��
	Do Until objRs.EOF Or i > objRs.PageSize'//���ǽ�
%>
	<TR onmouseover="this.style.backgroundColor='yellow';" onmouseout="this.style.backgroundColor='';">
		<TD><%=objRs("Num")%></TD>
		<TD><a href="./boardview.asp?Num=<%=objRs("Num")%>"><%=objRs("Title")%></a></TD>
		<TD><%=objRs("Name")%></TD>
		<TD><%=objRs("PostDate")%></TD>
		<TD><%=objRs("ReadCount")%></TD>
	</TR>
<%
		objRs.MoveNext()
		i = i + 1'//������
	Loop
%>
	<table>
<%
End Sub
%>
<form action="" method="get">
������ �̵� : <input type="text" name="Page" size=3>
<input type="submit" value="�̵�">
</form>
<form action="./boardsearchlist.asp" method="post">

�˻��� : <input type="text" name="SearchQuery">

<select name="SearchField">
	<option value="Title">����</option>
	<option value="Name">�̸�</option>
	<option value="Content">����</option>
</select>

<input type="submit" value="�˻�">

</form>

<br>

<a href="./boardwrite.asp">�۾���</a>



<hr>

<%
If Page > 1 Then
%>
	<a href="./boardlist.asp?Page=<%=Page-1%>">
	[����] 
	</a>
<%
Else
%>
	<a href="#" style="color:silver;">
	[����] 
	</a>
<%
End If
Response.Write("[" & Page & "/" & PageCount & "]")
If CInt(Page) < CInt(PageCount) Then
%>
	<a href="./boardlist.asp?Page=<%=Page+1%>">
	[����]
	</a>
<%
Else
%>
	<a href="#" style="color:silver;">
	[����]
	</a>
<%
End If
%>

<hr>

<%
Call AdvancedPaging(Page, PageCount)
%>

<%
Sub AdvancedPaging(PageNo, NumPage)'PageNo:����������, NumPage:��ü������
%>
	<!--���� 10��, ���� 10�� ����¡ ó�� ����-->
		<FONT style="font-size: 9pt;">
			<font color="#c0c0c0">[</font>
		<%	If PageNo > 10 Then %>
			<a href="./boardlist.asp?Page=<%= ((PageNo - 1) \ 10) * 10 %>">��</a>
		<%	Else %>
			<font color="#c0c0c0">��</font>
		<%	End If %>			
			<font color="#c0c0c0">|</font>
		<%	For i = (((PageNo - 1) \ 10) * 10 + 1) To ((((PageNo - 1) \ 10) + 1) * 10)
			If i > NumPage Then
				Exit For
			End If
			If i = Int(PageNo) Then
		%>
			<b><%= i %></b> <font color="#c0c0c0">|</font> 
		<%	Else %>
			<a href="./boardlist.asp?Page=<%= i %>"><%= i %></a> <font color="#c0c0c0">|</font> 
		<%	End If %>
		<%	Next %>
		<%	If CInt(i) < CInt(NumPage) Then %>
			<a href="./boardlist.asp?Page=<%= ((PageNo - 1) \ 10) * 10 + 11 %>">��</a>
		<%	Else %>
			<font color="#c0c0c0">��</font>
		<%	End If %>
			<font color="#c0c0c0">]</font>
		</FONT>
	<!--���� 10��, ���� 10�� ����¡ ó�� ����-->
<%
End Sub
%>

