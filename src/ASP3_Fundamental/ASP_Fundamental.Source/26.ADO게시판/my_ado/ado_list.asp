<%
'--------------------------------------------------
' Title : MyADO Education Version
' Program Name : ado_list.asp
' Program Description : Select���� �̿��� �Խ��� ����Ʈ �����ֱ�
' Include Files : adovbs.inc
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<!-- #include file="./scripts/adovbs.inc" -->
<%
'���� ����
'Option Explicit
Dim objRS   '���ڵ�� ��ü�� �ν��Ͻ�
Dim strSQL   '�����ͺ��̽� �˻� ���� ����
Dim strDB   '�����ͺ��̽� ���� ���ڿ� ����
Dim intPage   '���� �������� ��ȣ ����
Dim intTotalPage   '��ü �������� �� ����
Dim intCount   '�ӽ� ī���� ����

'�⺻ / �Ѱܿ� �������� ������ ����. ������ ù��° ������(1)�� ������.
intPage = Request.QueryString("intpage")
If intPage = "" Then
	intPage = 1
End If

'--------------------------------------------------------------------
'[1] ���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'[2] �����ͺ��̽� ���� ���ڿ� ����
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=ZZAZIPKIDOTCOM"

'[3] �����ͺ��̽� ó��(�˻�) ���� �ۼ�
strSQL = "Select * From my_ado Order By Num Desc"

'[4] ������������ ����
objRS.PageSize = 5

'[5] ���ڵ�� ����
objRs.Open strSQL, strDB, adOpenStatic, adLockReadOnly, adCmdtext
'--------------------------------------------------------------------
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

'�� �������� ���(���ڵ��� ���� ��Ÿ���� �Ӽ��� pagecount���
intTotalPage = objRS.PageCount 

'���� ����Ʈ ���������� ������ �������� ���������� ����
objRS.AbsolutePage = intPage

%>
		<DIV align="center">
			<TABLE border="1">
				<TR>
					<TD colspan="5" align="center" bgcolor="LightBlue">
						�ø��� ADO �Խ��Ǣ�
					</TD>
				</TR>
				<TR>
					<TD align="center">
						��ȣ
					</TD>
					<TD align="center">
						����
					</TD>
					<TD align="center">
						�ۼ���
					</TD>
					<TD align="center">
						�ۼ���
					</TD>
					<TD align="center">
						��ȸ
					</TD>
				</TR>
				<%
	intCount = 1   '1���� ������ ���ڵ� ����ŭ ���� �ݺ��� ���� ī��Ʈ����
	
	'�����ͺ��̽� ������ ���� ��� �Ǵ� ������ ���ڵ� ���� �ɶ����� �ݺ�.
	Do Until objRS.EOF OR intCount > objRS.PageSize   
	%>
				<TR>
					<TD align="center">
						<%=objRS.Fields("Num")%>
					</TD>
					<TD align="center">
						<A href="./ado_view.asp?num=<%=objRS.Fields("Num")%>&scrollaction=<%=intPage%>">
							<%=objRS.Fields("Title")%>
						</A>
					</TD>
					<TD align="center">
						<A href="mailto:<%=objRS.Fields("Email")%>">
							<%=objRS.Fields("Name")%>
						</A>
					</TD>
					<TD align="center">
						<%=objRS.Fields("PostDate")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("ReadCount")%>
					</TD>
				</TR>
				<%
	objRS.MoveNext
	'ī���ͺ����� ����	
	intCount = intCount + 1
	Loop
	%>
				<TR>
					<TD align="center" colspan="5">
						<%
	'���� ������ �̵� ��ũ �ۼ�
		If CInt(intPage) <> 1 Then
	%>
						<A href="ado_list.asp?intpage=<%=(intPage-1)%>">[����]</A>
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
						<A href="ado_list.asp?intpage=<%=(intPage+1)%>">[����]</A>
						<%
		Else
	%>
						<A href="#" style="color:gray;">[����]</A>
						<% 
		End If
	%>
					</TD>
				</TR>
				<TR>
					<TD align="center" colspan="5">
						<!-- -------------------------------- �˻��� ���� ------------------------------------- -->
						<DIV align="center">
							<TABLE border="0">
								<FORM name="ado_list_search_form" method="post" action="./ado_list_search_output.asp">
									<TR>
										<TD>
											�˻� : <SELECT name="strsearchfield" size="1">
												<OPTION value="name">
													�ۼ���</OPTION>
												<OPTION value="title">
													����</OPTION>
												<OPTION value="content">
													����</OPTION>
											</SELECT><INPUT type="text" name="strsearchquery" size="20"> <INPUT type="submit" value="�� ��">
										</TD>
									</TR>
								</FORM>
							</TABLE>
						</DIV>
						<!-- -------------------------------- �˻��� ���� ------------------------------------- -->
					</TD>
				</TR>
				<TR>
					<TD colspan="5" align="right">
						<A href="./ado_write_form.asp">[�۾���]</A>
						<% If Session("ID") <> "Admin" Then %>
						<A href="./ado_admin.asp">[������ �α���]</A>
						<% Else %>
						<A href="./ado_admin_logout.asp">[������ �α׾ƿ�]</A>
						<% End If %>
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
'���ڵ�� �ݱ�
objRS.Close

'���ڵ�� ��ü�� �ν��Ͻ� ����
Set objRS = Nothing
%>
