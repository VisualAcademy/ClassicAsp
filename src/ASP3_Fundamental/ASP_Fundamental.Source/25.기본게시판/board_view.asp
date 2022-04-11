<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_view.asp
' Program Description : �Խ��� ���� ���� �����ֱ�
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
Option Explicit

'���� ����
Dim objDB, objRS, strSQL, strSQL2, ScrollAction, Num

Num = Request.QueryString("num")

'���� ������� �̵��� ���� ����
ScrollAction = Request.QueryString("scrollaction")

'�����ͺ��̽�(connection��ü)�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")

'�����ͺ��̽� ����
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

'�ش� ��ȣ�� ��ȸ�� ������Ű�� ���� �ۼ�
strSQL = "Update my_board Set ReadCount = ReadCount + 1 "
strSQL = strSQL & "Where Num = " & Num
'�����
'Response.Write strSQL

'���� ����
objDB.Execute strSQL

'���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'�����ͺ��̽� ó��(�˻�) ���� �ۼ�
strSQL2 = "Select * From my_board Where Num = " & Num

'���ڵ�� ����(adOpenKeyset : ���ڵ���� ����, ���ڵ���� �������Ŀ� �ٸ�����ڰ� �߰��� �ڷḦ �� �� ����.)
'objRS.Open strSQL, objDB, adOpenKeyset
objRS.Open strSQL2, objDB, 1
%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
<%
If objRS.BOF OR objRS.EOF Then
%>
		<p align="center">
			�Էµ� �����Ͱ� �����ϴ�.
		</p>
<%
Else
%>
		<div align="center">
			<table border="1">
				<tr>
					<td colspan="2" align="center">
						�ø��� �Խ��� �ۺ����
					</td>
				</tr>
				<tr>
					<td align="center">����</td>
					<td align="center"><%=objRS.Fields("Title")%></td>
				</tr>
				<tr>
					<td align="center">�ۼ���</td><td align="center"><%=objRS.Fields("Name")%></td>
				</tr>
				<tr>
					<td align="center">�ۼ���</td><td><%=objRS.Fields("PostDate")%></td>
				</tr>
				<tr>
					<td align="center">��ȸ</td><td><%=objRS.Fields("ReadCount")%></td>
				</tr>
				<tr>
					<td align="center" colspan="2"><%=Replace(objRS.Fields("Content"),Chr(13),"<br>")%></td>
				</tr>
				<tr>
					<td colspan="2" align="right">
						<a href="./board_write_form.asp">[�۾���]</a>&nbsp;
						<a href="./board_modify.asp?num=<%=objRS.Fields("Num")%>">[����]</a>&nbsp;
						<a href="./board_delete.asp?num=<%=objRS.Fields("Num")%>">[����]</a>&nbsp;
						<a href="./board_list.asp?intpage=<%=ScrollAction%>">[����Ʈ]</a>
					</td>
				</tr>
				<% If Session("ID") = "Admin" Then %>
				<tr>
					<td colspan="2" align="center">
						<a href="./board_delete_process.asp?num=<%=objRS.Fields("Num")%>&passwd=<%=objRS.Fields("Passwd")%>">����� �����(<%=objRS.Fields("Passwd")%>)</a>
					</td>
				</tr>
				<% End If %>
			</table>
		</div>
<%
End If
%>
	</BODY>
</HTML>
<%
'���ڵ�� �ݱ�
objRS.Close

'�����ͺ��̽� �ݱ�
objDB.Close

'���ڵ�� ��ü�� �ν��Ͻ� ����
Set objRS = Nothing

'�����ͺ��̽� ��ü�� �ν��Ͻ� ����
Set objDB = Nothing
%>
