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
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

'���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'�����ͺ��̽� ó��(�˻�) ���� �ۼ�
strSQL = "Select * From my_memo Order By Num Desc"

'���ڵ�� ����
objRS.Open strSQL, objDB, 1

%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
If objRS.BOF OR objRS.EOF Then
%>
		<P align="center">
			�Էµ� �����Ͱ� �����ϴ�.
		</P>
		<%
Else
%>
		<DIV align="center">
			<TABLE border="1">
				<TR>
					<TD align="center">
						�� ȣ
					</TD>
					<TD align="center">
						�� ��
					</TD>
					<TD align="center">
						�� �� ��
					</TD>
					<TD align="center">
						��&nbsp;&nbsp;&nbsp;��
					</TD>
					<TD align="center">
						�� ¥
					</TD>
				</TR>
				<%
	Do Until objRS.EOF   '�����ͺ��̽� ������ ���� ���.
	%>
				<TR>
					<TD align="center">
						<%'=objRS.Fields("Num")%>
						<%'=objRS.Fields(0)%>
						<%=objRS("Num")%>
						<%'=objRS(0)%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("Name")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("Email")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("Title")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("PostDate")%>
					</TD>
				</TR>
				<%
	objRS.MoveNext   '���ڵ带 ���� ���ڵ�� �̵���Ű�� ���ڵ�� �޼ҵ�
	Loop
	%>
			</TABLE>
		</DIV>
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
