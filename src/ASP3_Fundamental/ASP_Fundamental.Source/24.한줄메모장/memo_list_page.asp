<%
'--------------------------------------------------------------------
' Title : MyMemo Education Version
' Program Name : memo_list_page.asp
' Program Description : ����¡  �� �˻��� ó���� ����Ʈ
' Include Files : None
' Last Update : 
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.zzazipki.com/
'--------------------------------------------------------------------
%>
<%
Option Explicit

'���� ����
Dim objDB, objRS, strSQL, intPage, intTotalPage, intCount

'--------------------------------------------------------------------
'�⺻ / �Ѱܿ� �������� ������ ����. ������ ù��° ������(1)�� ������.
intPage = Request.QueryString("intpage")
IF intPage = "" Then
	intPage = 1
End IF
'--------------------------------------------------------------------

'�����ͺ��̽�(connection��ü)�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")

'�����ͺ��̽� ����
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'�����ͺ��̽� ó��(�˻�) ���� �ۼ�
strSQL = "Select * From my_memo Order By Num Desc"

'--------------------------------------------------------------------
'RecordSet ��ü�� PageSize �Ӽ� ��� : �� �������� ������ ����
objRS.PageSize = 5
'--------------------------------------------------------------------

'���ڵ�� ����(sql����, db���Ṯ�ڿ�, Ŀ��Ÿ��, ��Ÿ��, �ɼ�)
'objRS.Open strSQL, objDB,adOpenKeyset
'����¡ ó���� ���ڵ���� Ŀ��Ÿ���� �ݵ�� 1 �Ǵ� 3�� ���� �־����.
objRS.Open strSQL, objDB, 1

%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
IF objRS.Bof or objRS.Eof then
%>
		<P align="center">
			�Էµ� �����Ͱ� �����ϴ�.
		</P>
		<%
Else

'---------------------------------------------------------------------
'�� �������� ���(���ڵ��� ���� ��Ÿ���� �Ӽ��� pagecount���)
intTotalPage = objRS.PageCount 

'���� ����Ʈ ���������� ������ �������� ���������� ����(Ư�� �������� �̵�)
objRS.AbsolutePage = intPage
'---------------------------------------------------------------------

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
						��&nbsp;&nbsp;&nbsp;��
					</TD>
					<TD align="center">
						�� ¥
					</TD>
				</TR>
				<%
	'---------------------------------------------------------------------
	intCount = 1   '1���� ������ ���ڵ� ����ŭ ���� �ݺ��� ���� ī��Ʈ����
	
	'�����ͺ��̽� ������ ���� ��� �Ǵ� ������ ���ڵ� ���� �ɶ����� �ݺ�.
	Do Until objRS.EOF OR intCount > objRS.PageSize   
	'---------------------------------------------------------------------	
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
	'---------------------------------------------------------------------	
	'ī���ͺ����� ����	
	intCount = intCount + 1
	'---------------------------------------------------------------------		
	Loop
	%>
				<TR>
					<TD align="center" colspan="4">
						<%
	'---------------------------------------------------------------------			
	'���� ������ �̵� ��ũ �ۼ�
		IF CInt(intPage) <> 1 Then
	%>
						<A href="memo_list_page.asp?intpage=<%=(intPage-1)%>">[����]</A>
						<%
		Else
	%>
						<A href="#" style="color:gray;">[����]</A>
						<% 
		End IF
	%>
						<%
	'���� ������ �� ��ü ������ ���.
	
	Response.Write("[����:" & intPage & "/")
	Response.Write("��ü:" & intTotalPage & "]")	
	
	%>
						<%
	'���� ������ �̵� ��ũ �ۼ�
		IF CInt(intPage) <> CInt(intTotalPage) then
	%>
						<A href="memo_list_page.asp?intpage=<%=(intPage+1)%>">[����]</A>
						<%
		Else
	%>
						<A href="#" style="color:gray;">[����]</A>
						<% 
		End IF
	'---------------------------------------------------------------------				
	%>
					</TD>
				</TR>
			</TABLE>
		</DIV>
		<%
End IF
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
