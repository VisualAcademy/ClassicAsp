<%
'--------------------------------------------------
' Title : Myupload Education Version
' Program Name : upload_list_search_output.asp
' Program Description : �Խ��� �˻� ��� �����ֱ�
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
Option Explicit
'---------------------------------------------------------------------
'���� ����
Dim strSearchField, strSearchQuery, objDB, objRS, strSQL, intTotalPage, intPage, intCount
'---------------------------------------------------------------------
'�⺻ / �Ѱܿ� �������� ������ ����. ������ ù��° ������(1)�� ������.
intPage = Request.QueryString("intpage")
IF intPage = "" Then
	intPage = 1
End IF

'---------------------------------------------------------------------
'�Ѱ����� �˻�Ű���� 2���� ���� ������ ����.
strSearchField = Request("strsearchfield")
strSearchQuery = Request("strsearchquery")
'---------------------------------------------------------------------

'Connection��ü�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")

'�����ͺ��̽� ����
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")
         
'���ڵ� �� ��ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet")

'---------------------------------------------------------------------
'SQL Query ����
strSQL = "Select * From my_upload Where " & strSearchField & " Like '%" & _
         strSearchQuery & "%' Order By Num Desc"
'---------------------------------------------------------------------

'���ڵ���� ������ ����Ʈ ����
objRS.PageSize = 5

'���ڵ� �� ����
objRS.Open strSQL, objDB, 1
%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
<%
IF objRS.BOF OR objRS.EOF Then
'---------------------------------------------------------------------
%>
		<p align="center">
			<%=strSearchQuery%>
			�� �˻��� ������ ���, �˻��� �����Ͱ� �����ϴ�.
		</p>
		<%
'---------------------------------------------------------------------		
Else

'�� �������� ���(���ڵ��� ���� ��Ÿ���� �Ӽ��� pagecount���
intTotalPage = objRS.PageCount 

'���� ����Ʈ ���������� ������ �������� ���������� ����
objRS.AbsolutePage = intPage
%>
		<div align="center">
			<table border="1">
				<tr>
					<td colspan="5" align="center">
						�� �� �� ��
					</td>
				</tr>
				<tr>
					<td align="center">��ȣ</td>
					<td align="center">����</td>
					<td align="center">�ۼ���</td>
					<td align="center">�ۼ���</td>
					<td align="center">��ȸ</td>
				</tr>
	<%
	intCount = 1   '1���� ������ ���ڵ� ����ŭ ���� �ݺ��� ���� ī��Ʈ����
	
	'�����ͺ��̽� ������ ���� ��� �Ǵ� ������ ���ڵ� ���� �ɶ����� �ݺ�.
	Do Until objRS.EOF OR intCount > objRS.PageSize   
	%>
				<tr>
					<td align="center"><%=objRS.Fields("Num")%></td>
					<td align="center"><a href="./upload_view.asp?num=<%=objRS.Fields("Num")%>"><%=objRS.Fields("Title")%></a></td>
					<td align="center"><a href="mailto:<%=objRS.Fields("Email")%>"><%=objRS.Fields("Name")%></a></td>
					<td align="center"><%=objRS.Fields("PostDate")%></td>
					<td align="center"><%=objRS.Fields("ReadCount")%></td>
				</tr>
	<%
	objRS.MoveNext
	'ī���ͺ����� ����	
	intCount = intCount + 1
	loop
	%>
				<tr>
					<td align="center" colspan="5">
	<%

	'���� ������ �̵� ��ũ �ۼ�
		IF CInt(intPage) <> 1 Then
	%>
						<a href="upload_list_search_output.asp?intpage=<%=(intPage-1)%>&strsearchfield=<%=strSearchField%>&strsearchquery=<%=strSearchQuery%>">[����]</a>
	<%
		Else
	%>
						<a href="#" style="color:gray;">[����]</a>
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
		IF CInt(intPage) <> CInt(intTotalPage) Then
	%>
						<a href="upload_list_search_output.asp?intpage=<%=(intPage+1)%>&strsearchfield=<%=strSearchField%>&strsearchquery=<%=strSearchQuery%>">[����]</a>
	<%
		Else
	%>
						<a href="#" style="color:gray;">[����]</a>
	<% 
		End IF
				'---------------------------------------------------------------------
	%>
					</td>
				</tr>
				<tr>
					<td align="center" colspan="5">
						<a href="./upload_list.asp">[�˻��Ϸ�]</a>
					</td>
				</tr>
			</table>
		</div>
		<%
				'---------------------------------------------------------------------		
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
