<%
'--------------------------------------------------
' Title : Basic ����
' Program Name : boardview.asp
' Program Description : ���� �ۺ��� ������
' QueryString : �ݵ�� list.asp���� Num=1 ������ ����
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
'[1] ���� ����
Dim Num: Num = Request("Num")
Dim objCon: Dim objRs: Dim strSql
'[2] Ŀ�ؼ� �ν��Ͻ�
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] ����
objCon.Open(Application("ConnectionString"))
'[!] ��ȸ�� ���� : �ѹ��� üũ -> ��Ű ����
If Request.Cookies("ReadCount")(Num) = "" Then 
  strSql = "Update Basic Set ReadCount = ReadCount + 1 Where Num = " & Num
  objCon.Execute(strSql)
  Response.Cookies("ReadCount")(Num) = "OK"
End If
'[4] ���ڵ�� �ν��Ͻ�
Set objRs = Server.CreateObject("ADODB.RecordSet")
'[5] ���ڵ�� ���� : ��� ����(select)
strSql = "Select * From Basic Where Num = " & Num
objRs.Open strSql, objCon
'[6] ���
If objRs.BOF Or objRs.EOF Then
	Response.Write("�ش� ���ڵ尡 �������� �ʽ��ϴ�.")
Else 'HTML+ASP�ڵ�(���İ�Ƽ�ڵ�)�� ��� ����� ���
	Call ShowRecordSet(objRs)
End If
'[7] �ݱ�
objRs.Close(): objCon.Close()
'[8] ����
Set objRs = Nothing: Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>

	<center><h3>�⺻ �Խ��� ���� �׸� ����</h3></center>
	
	<center>
	<TABLE border=1 width=100%>
	<TR>
		<TD>�۾���</TD>	<TD><%=objRs("Name")%></TD>
		<TD>��¥</TD>	<TD><%=objRs("PostDate")%></TD>
	</TR>
	<TR>
		<TD>Email</TD><TD><%=objRs("Email")%></TD>
		<TD>��ȸ��</TD>	<TD><%=objRs("ReadCount")%></TD>
	</TR>
	<TR>
		<TD>����</TD>	<TD colspan=3><%=objRs("Title")%></TD>
	</TR>
	<TR>
		<TD colspan=4><%=objRs("Content")%></TD>
	</TR>
	</TABLE>
	<a href="./boardlist.asp">[�Խ��Ǹ���Ʈ]</a>
	<a href="./boardmodify.asp?Num=<%=objRs("Num")%>">[����]</a>
	<a href="./boarddelete.asp?Num=<%=objRs("Num")%>">[����]</a>
	</center>
<%
End Sub
%>