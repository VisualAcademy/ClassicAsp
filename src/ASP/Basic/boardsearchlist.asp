<%
'--------------------------------------------------
' Title : Basic ����
' Program Name : boardsearchlist.asp
' Program Description : �˻� ��� ����Ʈ ������
' QueryString : �˻������� SearchField, SearchQuery
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
'[1] ���� ����
Option Explicit
Dim SearchField: SearchField = Request("SearchField")
Dim SearchQuery: SearchQuery = Request("SearchQuery")
Dim objCon: Dim objRs: Dim strSql
'[2] Ŀ�ؼ� ��ü�� �ν��Ͻ�
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] Open() : Ŀ�ؼ� ��Ʈ�� ����
objCon.Open(Application("ConnectionString"))
'[4] ���ڵ�� ��ü�� �ν��Ͻ�
Set objRs = Server.CreateObject("ADODB.RecordSet")
'[5] Open() : ��� ����(Select)
strSql = "Select * From Basic Where " & _
		 SearchField & " Like '%" & _
		 SearchQuery & "%' Order By Num Desc"
objRs.Open strSql, objCon
'[6] ���
If objRs.BOF Or objRs.EOF Then
	Response.Write("<b style='color:red;'>" & SearchQuery & "</b>�� �˻��� ����� �����ϴ�.")
Else
	Call ShowRecordSet(objRs)
End If
'[7] Close()
objRs.Close(): objCon.Close()
'[8] ��ü ����
Set objRs = Nothing: Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>
	<table border=1 width="100%">
<%
	While Not objRs.EOF
%>
	<TR onmouseover="this.style.backgroundColor='red';" onmouseout="this.style.backgroundColor='';">
		<TD><%=objRs("Num")%></TD>
		<TD><a href="./boardview.asp?Num=<%=objRs("Num")%>"><%=objRs("Title")%></a></TD>
		<TD><%=objRs("Name")%></TD>
		<TD><%=objRs("Email")%></TD>
		<TD><%=objRs("ReadCount")%></TD>
	</TR>
<%
		objRs.MoveNext()
	WEnd
%>
	<table>
<%
End Sub
%>
<a href="./boardlist.asp">�˻�����</a>