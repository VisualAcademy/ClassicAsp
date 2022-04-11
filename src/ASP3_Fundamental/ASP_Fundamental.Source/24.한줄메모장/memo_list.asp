<%
'--------------------------------------------------------------------
' Title : MyMemo Education Version
' Program Name : memo_list.asp
' Program Description : �޸�� ����Ʈ ������
' Include Files : None
' Last Update : 2001.08.26.
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<%
'[1] ������ ����� ����
Option Explicit

'[2] ���� ����
Dim objDB			'Connection��ü�� �ν��Ͻ� ��ü�� ������ ����
Dim objRS			'RecordSet��ü�� �ν��Ͻ� ��ü�� ������ ���� 
Dim strSQL			'MY�޸��� DB���� �����͸� �˻��ϴ� SQL���� �����ϴ� ����		
Dim intPage			'MY�޸��� ���� ������ ��ȣ�� ������ ����
Dim intTotalPage	'MY�޸��� ��ü ������ ��ȣ�� ������ ���� 
Dim intCount		'1���� ������ ���ڵ� ����ŭ ���� �ݺ��� ���� ī��Ʈ����

'--------------------------------------------------------------------
'[3] ������ �� ���� : ���� ���۵� ���������� ������ ���� ������ ������ �����ϰ�, 
'�׷��� ������ ù��° ������(1)�� ������.
intPage = Request.QueryString("intPage")
If intPage = "" Then
	intPage = 1
End if
'--------------------------------------------------------------------

'[4] �����ͺ��̽�(connection��ü)�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")

'[5] �����ͺ��̽� ����(ODBC������ DSN������ SQL Server�� �α��� ����(uid/pwd) ���)
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'[6] ���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'[7] �����ͺ��̽� ó��(�˻�) ���� �ۼ�
strSQL = "Select * From my_memo Order By Num DESC"

'--------------------------------------------------------------------
'[8] RecordSet ��ü�� PageSize �Ӽ� ��� : �� �������� ������ ����
objRS.PageSize = 5
'--------------------------------------------------------------------

'[9] ���ڵ�� ����
'���ڵ�� ����(sql����, db���Ṯ�ڿ�, Ŀ��Ÿ��, ��Ÿ��, �ɼ�)
'ADO��� ����(objRS.Open strSQL, objDB, adOpenStatic, adLockReadOnly, adCmdtext)
'����¡ ó���� ���ڵ���� Ŀ��Ÿ���� �ݵ�� 1 �Ǵ� 3�� ���� �־����.
objRS.Open strSQL, objDB, 1

'[10] MY�޸��� ����Ʈ ������ ó�� �κ� ����
%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
'[11] MY �޸��� DB�� �����Ͱ� ���� ��� ���� �ۼ�
If objRS.BOF OR objRS.EOF Then
%>
		<P align="center">
			�Էµ� �����Ͱ� �����ϴ�.
		</P>
		<P align="center">
			<A href="./memo_write_form.asp">[�۾���]</A>
		</P>
		<%
Else

'--------------------------------------------------------------------
'[12] �� �������� ���(���ڵ��� ���� ��Ÿ���� �Ӽ��� PageCount���)
intTotalPage = objRS.PageCount 

'[13] ���� ����Ʈ ���������� ������ �������� ���������� ����(Ư�� �������� �̵�)
objRS.AbsolutePage = intPage
'--------------------------------------------------------------------
%>
		<DIV align="center">
			<TABLE border="1" width="100%">
				<TR>
					<TD colspan="5" align="center">
						<nobr>
						��MY �޸����
						</nobr>
					</TD>
				</TR>
				<TR>
					<TD align="center">
						�� ȣ
					</TD>
					<TD align="center">
						�� ��
					</TD>
					<TD align="center">
						��&nbsp;&nbsp;&nbsp;��
					</TD>
					<TD align="center">
						�� ¥
					</TD>
				</TR>
				<%
'[14] MY �޸��� ����Ʈ ó��
	intCount = 1   '1���� ������ ���ڵ� ����ŭ ���� �ݺ��� ���� ī��Ʈ����
	
	'�����ͺ��̽� ������ ���� ��� �Ǵ� ������ ���ڵ� ���� �ɶ����� �ݺ�.
	Do Until objRS.EOF OR intCount > objRS.PageSize   
	%>
				<TR>
					<TD align="center">
						<%=objRS.Fields("Num")%>
					</TD>
					<TD align="center">
						<A href="mailto:<%=objRS.Fields("Email")%>">
							<%=objRS.Fields("Name")%>
						</A>
					</TD>
					<TD align="center">
						<%=objRS.Fields("Title")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("PostDate")%>
					</TD>
				</TR>
				<%
	objRS.MoveNext
	'ī���ͺ����� ����	
	intCount = intCount + 1
	Loop
	%>
				<TR>
					<TD colspan="5" align="right">
						<A href="./memo_write_form.asp">[�۾���]</A>&nbsp;&nbsp; 
						<A href="./memo_delete.asp">[�����]]</A>
					</TD>
				</TR>
				<TR>
					<TD align="center" colspan="4">
						<%
'--------------------------------------------------------------------
'[15] ������ ó��
	'���� ������ �̵� ��ũ �ۼ�
		If CInt(intPage) <> 1 Then
	%>
						<A href="memo_list.asp?intpage=<%=(intPage-1)%>">[����]</A>
						<%
		Else
	%>
						<A href="#" style="color:gray;">[����]</A>
						<% 
		End If
	%>
						<%
	'���� ������ �� ��ü ������ ���.
	
	Response.Write("[����:" & intPage & "/")
	Response.Write("��ü:" & intTotalPage & "]")	
	
	%>
						<%
	'���� ������ �̵� ��ũ �ۼ�
		If CInt(intPage) <> CInt(intTotalPage) Then
	%>
						<A href="memo_list.asp?intpage=<%=(intPage+1)%>">[����]</A>
						<%
		Else
	%>
						<A href="#" style="color:gray;">[����]</A>
						<% 
		End If
		'--------------------------------------------------------------------

'[16] �˻��� ����
	%>
					</TD>
				</TR>
				<TR>
					<TD align="center" colspan="4">
						<!-- -------------------------------- �˻��� ���� ------------------------------------- -->
						<DIV align="center">
							<TABLE border="0">
								<FORM name="memo_list_search_form" method="post" action="./memo_list_search_output.asp">
									<TR>
										<TD>
											�˻� : <SELECT name="strsearchfield" size="1">
												<OPTION value="name">
													�̸�</OPTION>
												<OPTION value="title">
													�޸�</OPTION>
											</SELECT><INPUT type="text" name="strsearchquery" size="20"> <INPUT type="submit" value="�� ��">
										</TD>
									</TR>
								</FORM>
							</TABLE>
						</DIV>
						<!-- -------------------------------- �˻��� ���� ------------------------------------- -->
					</TD>
				</TR>
			</TABLE>
		</DIV>
		<%
End If
%>
	</BODY>
</HTML>
<%
'[17] ���ڵ�� �ݱ�
objRS.Close

'[18] �����ͺ��̽� �ݱ�
objDB.Close

'[19] ���ڵ�� ��ü�� �ν��Ͻ� ����
Set objRS = Nothing

'[20] �����ͺ��̽� ��ü�� �ν��Ͻ� ����
Set objDB = Nothing
%>
