<%
'--------------------------------------------------
' Title : Myado Education Version
' Program Name : ado_view.asp
' Program Description : �Խ��� ���� ���� �����ֱ�
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!--METADATA TYPE= "typelib" NAME= "ADODB Type Library" FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll" -->

<%
Option Explicit

'���� ����
Dim objDB, objRS, strSQL, strSQL2, ScrollAction, Num
Dim strDB

Num = Request("num")

'���� ������� �̵��� ���� ����
ScrollAction = Request.QueryString("scrollaction")

'--------------------------------------------------------------------
'[1] ���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'[2] �����ͺ��̽� ���� ���ڿ� ����
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=ZZAZIPKIDOTCOM"

'[3] �Ѱܿ� �۹�ȣ�� �ش�Ǵ� �н����� ��������
strSQL = "Select * From my_ado Where Num = " & Num
'Response.Write strSQL

'[4] ���ڵ�� ����
objRs.Open strSQL, strDB, adOpenStatic, adLockPessimistic, adCmdText

With objRS
	
		.Fields("ReadCount") = .Fields("ReadCount") + 1   '��ȸ�� 1 ����

		.Update

		.Close

End With

Set objRS = Nothing
'--------------------------------------------------------------------

'--------------------------------------------------------------------
'���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'�����ͺ��̽� ó��(�˻�) ���� �ۼ�
strSQL2 = "Select * From my_ado Where Num = " & Num

'���ڵ�� ����
objRs.Open strSQL2, strDB, adOpenStatic, adLockPessimistic, adCmdText
'--------------------------------------------------------------------
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
						<a href="./ado_write_form.asp">[�۾���]</a>&nbsp;
						<a href="./ado_modify.asp?num=<%=objRS.Fields("Num")%>">[����]</a>&nbsp;
						<a href="./ado_delete.asp?num=<%=objRS.Fields("Num")%>">[����]</a>&nbsp;
						<a href="./ado_list.asp?intpage=<%=ScrollAction%>">[����Ʈ]</a>
					</td>
				</tr>
				<% If Session("ID") = "Admin" Then %>
				<tr>
					<td colspan="2" align="center">
						<a href="./ado_delete_process.asp?num=<%=objRS.Fields("Num")%>&passwd=<%=objRS.Fields("Passwd")%>">����� �����(<%=objRS.Fields("Passwd")%>)</a>
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

'���ڵ�� ��ü�� �ν��Ͻ� ����
Set objRS = Nothing

%>
