<%
'--------------------------------------
' MyMemo Education Version
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------
%>
<%
Option Explicit

'���� ����
Dim objDB, objRS, strSQL

'�����ͺ��̽�(connection��ü)�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")

'�����ͺ��̽� ����
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'�����ͺ��̽� ó��(�˻�) ���� �ۼ�
strSQL = "Select * From my_memo Order By Num Desc"

'���ڵ�� ����
objRS.Open strSQL, objDB, 1

%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
If objRS.BOF or objRS.EOF then
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
					<td align="center">
						�� ȣ
					</td>
					<td align="center">
						�� ��
					</td>
					<td align="center">
						��&nbsp;&nbsp;&nbsp;��
					</td>
					<td align="center">
						�� ¥
					</td>
				</tr>
				<%
	Do Until objRS.EOF   '�����ͺ��̽� ������ ���� ���.
	%>
				<tr>
					<td align="center">
						<%=objRS.Fields("Num")%>
					</td>
					<td align="center">
						<a href="mailto:<%=objRS.Fields("Email")%>">
							<%=objRS.Fields("Name")%>
						</a>
					</td>
					<td align="center">
						<%=objRS.Fields("Title")%>
					</td>
					<td align="center">
						<%=objRS.Fields("PostDate")%>
					</td>
				</tr>
				<%
	objRS.MoveNext
	Loop
	%>
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
