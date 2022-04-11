<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_list.asp
' Program Description : Select���� �̿��� �Խ��� ����Ʈ �����ֱ�
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
Option Explicit

'���� ����
Dim objDB, objRS, strSQL, intPage, intTotalPage, intCount

'�⺻ / �Ѱܿ� �������� ������ ����. ������ ù��° ������(1)�� ������.
intPage = Request.QueryString("intpage")
If intPage = "" Then
	intPage = 1
End If

'�����ͺ��̽�(connection��ü)�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")

'�����ͺ��̽� ����
objDB.open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

'���ڵ�°�ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'�����ͺ��̽� ó��(�˻�) ���� �ۼ�
strSQL = "Select * From my_board Order By Num Desc"

'������������ ����
objRS.PageSize = 5

'���ڵ�� ����(adOpenKeyset : ���ڵ���� ����, ���ڵ���� �������Ŀ� �ٸ�����ڰ� �߰��� �ڷḦ �� �� ����.)
'objRS.Open strSQL, objDB, adOpenKeyset
objRS.Open strSQL, objDB, 1
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

'�� �������� ���(���ڵ��� ���� ��Ÿ���� �Ӽ��� pagecount���
intTotalPage = objRS.PageCount 

'���� ����Ʈ ���������� ������ �������� ���������� ����
objRS.AbsolutePage = intPage

%>
		<div align="center">
			<table border="1">
				<tr>
					<td colspan="5" align="center">
						�ø��� �Խ��Ǣ�
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
					<td align="center"><a href="./board_view.asp?num=<%=objRS.Fields("Num")%>&scrollaction=<%=intPage%>"><%=objRS.Fields("Title")%></a></td>
					<td align="center"><a href="mailto:<%=objRS.Fields("Email")%>"><%=objRS.Fields("Name")%></a></td>
					<td align="center"><%=objRS.Fields("PostDate")%></td>
					<td align="center"><%=objRS.Fields("ReadCount")%></td>
				</tr>
	<%
	objRS.MoveNext
	'ī���ͺ����� ����	
	intCount = intCount + 1
	Loop
	%>
				<tr>
					<td align="center" colspan="5">
	<%
	'���� ������ �̵� ��ũ �ۼ�
		If CInt(intPage) <> 1 Then
	%>
		<a href="board_list.asp?intpage=<%=(intPage-1)%>">[����]</a>
	<%
		Else
	%>
		<a href="#" style="color:gray;">[����]</a>
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
		<a href="board_list.asp?intpage=<%=(intPage+1)%>">[����]</a>
	<%
		Else
	%>
						<a href="#" style="color:gray;">[����]</a>
	<% 
		End If
	%>
					</td>
				</tr>
				<tr>
					<td align="center" colspan="5">
						<!-- -------------------------------- �˻��� ���� ------------------------------------- -->
						<div align="center">
							<table border="0">
								<form name="board_list_search_form" method="post" action="./board_list_search_output.asp">
									<tr>
										<td>
											�˻� : <select name="strsearchfield" size="1">
												<option value="name">
													�ۼ���</option>
												<option value="title">
													����</option>
												<option value="content">
													����</option>
											</select><input type="text" name="strsearchquery" size="20"> <input type="submit" value="�� ��">
										</td>
									</tr>
								</form>
							</table>
						</div>
						<!-- -------------------------------- �˻��� ���� ------------------------------------- -->
					</td>
				</tr>
				<tr>
					<td colspan="5" align="right">
						<a href="./board_write_form.asp">[�۾���]</a>

						<% If Session("ID") <> "Admin" Then %>
						<a href="./board_admin.asp">[������ �α���]</a>
						<% Else %>
						<a href="./board_admin_logout.asp">[������ �α׾ƿ�]</a>
						<% End If %>

					</td>
				</tr>
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
