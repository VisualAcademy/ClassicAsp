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
'---------------------------------------------------------------------
'���� ����
Dim strSearchField, strSearchQuery, objDB, objRS, strSQL, intTotalPage, intPage, intCount
'---------------------------------------------------------------------
'�⺻ / �Ѱܿ� �������� ������ ����. ������ ù��° ������(1)�� ������.
intPage = Request.QueryString("intpage")
If intPage = "" then
	intPage = 1
End If

'---------------------------------------------------------------------
'�Ѱ����� 2���� ���� ������ ����.
strSearchField = Request("strsearchfield")
strSearchQuery = Request("strsearchquery")
'---------------------------------------------------------------------

'Connection��ü�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

         
'���ڵ� �� ��ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet")

'---------------------------------------------------------------------
'SQL Query ����
strSQL = "Select * From my_memo Where " & strSearchField & " Like '%" & _
         strSearchQuery & "%' Order By Num Desc"
'---------------------------------------------------------------------

objRS.PageSize = 5

'���ڵ� �� ����
objRS.Open strSQL, objDB, 1

%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
If objRS.BOF or objRS.EOF then
'---------------------------------------------------------------------
%>
		<P align="center">
			<font color="red"><%=strSearchQuery%></font>
			�� �˻��� ������ ���, �˻��� �����Ͱ� �����ϴ�.
		</P>
		<P align="center">
			<input type="button" value="�ڷΰ���" onclick="history.go(-1);">
		</P>
		<%
'---------------------------------------------------------------------		
Else

'�� �������� ���(���ڵ��� ���� ��Ÿ���� �Ӽ��� pagecount���
intTotalPage = objRS.PageCount 

'���� ����Ʈ ���������� ������ �������� ���������� ����
objRS.AbsolutePage = intPage

%>
		<DIV align="center">
			<TABLE border="1" width="100%">
				<TR>
					<TD align="center" colspan="4">
						�� �޸� �˻��ϱ��
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
	intCount = 1   '1���� ������ ���ڵ� ����ŭ ���� �ݺ��� ���� ī��Ʈ����
	
	'�����ͺ��̽� ������ ���� ��� �Ǵ� ������ ���ڵ� ���� �ɶ����� �ݺ�.
	Do Until objRS.EOF OR intCount > objRS.pagesize   
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
					<TD align="center" colspan="4">
						<%

	'���� ������ �̵� ��ũ �ۼ�
		If CInt(intPage) <> 1 then
	%>
						<A href="./memo_list_search_output.asp?intpage=<%=(intPage-1)%>&strsearchfield=<%=strSearchField%>&strsearchquery=<%=strSearchQuery%>">
							[����]</A>
						<%
		else
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
		If CInt(intPage) <> CInt(intTotalPage) then
	%>
						<A href="./memo_list_search_output.asp?intpage=<%=(intPage+1)%>&strsearchfield=<%=strSearchField%>&strsearchquery=<%=strSearchQuery%>">
							[����]</A>
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
					<TD align="center" colspan="4">
						<A href="./memo_list.asp">[�˻��Ϸ�]</A>
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

'�����ͺ��̽� �ݱ�
objDB.Close

'���ڵ�� ��ü�� �ν��Ͻ� ����
Set objRS = Nothing

'�����ͺ��̽� ��ü�� �ν��Ͻ� ����
Set objDB = Nothing
%>
