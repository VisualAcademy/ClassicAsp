<%
Dim objCon: Dim objCmd: Dim objRs
Dim intNum: intNum = Request("Num")
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
	Set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = objCon
	'//��ȸ�� ����
	objCmd.CommandText = "Update Basic Set ReadCount = ReadCount + 1 Where Num = " & intNum
	objCmd.Execute()
	'//��ȸ�� ���� ��
	' ���������� -> " & ���� & "
	objCmd.CommandText = "Select * From Basic Where Num = " & intNum 
		Set objRs = objCmd.Execute()
			Call ShowRecordSet(objRs)
		objRs.Close()
		Set objRs = Nothing
	Set objCmd = Nothing
objCon.Close()
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)'// ��� ��� ���� �Լ�
			If objRs.BOF Or objRs.EOF Then
			Else
				Do Until objRs.EOF
%>	
	<TABLE border="1" width="600" align="center">
	<TR>
		<TD>��ȣ : </TD><TD><%=objRs.Fields("Num")%></TD>
	</TR>
	<TR>
		<TD>���� : </TD><TD><%=objRs.Fields("Title")%></TD>
	</TR>
	<TR>
		<TD>�ۼ��� : </TD><TD><%=objRs.Fields("Name")%></TD>
	</TR>
	<TR>
		<TD>�ۼ��� : </TD><TD><%=objRs.Fields("PostDate")%></TD>
	</TR>
	<TR>
		<TD>��ȸ�� : </TD><TD><%=objRs.Fields("ReadCount")%></TD>
	</TR>
	<TR>
		<TD colspan="2" height="200" valign="top">
<%	
	Dim strContent
	'[!] HTML �±� ó��
	strContent = objRs.Fields("Content")
	If objRs.Fields("Encoding") = "Text" Then ' Text
		strContent = Replace(strContent, "&", "&amp;")
		strContent = Replace(strContent, "<", "&lt;")
		strContent = Replace(strContent, ">", "&gt;")
		'[!] ���� ó��
		strContent = Replace(strContent,Chr(13) & Chr(10),"<br>")
	ElseIf objRs("Encoding") = "Mixed" Then ' Mixed
		'[!] ���� ó��
		strContent = Replace(strContent,Chr(13) & Chr(10),"<br>")
	End If
	Response.Write(strContent)
%>
		</TD>
	</TR>
	<TR>
		<TD colspan="2" align="right">
			<a href="./BoardModify.asp?Num=<%=objRs("Num")%>">[����]</a>
			<a href="./BoardDelete.asp?Num=<%=objRs("Num")%>">[����]</a>
			<a href="./BoardList.asp">[����Ʈ]</a>
		</TD>
	</TR>
	</TABLE>				
<%
					objRs.MoveNext()
				Loop
			End If
End Sub
%>
