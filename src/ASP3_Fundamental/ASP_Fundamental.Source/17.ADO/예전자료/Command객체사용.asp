<%
'--------------------------------------------------------------------
' Title : ADO - Command��ü
' Program Name : Command��ü���.asp
' Program Description : Command��ü�� ��� ��� ���
' Include Files : None
' Copyright (C) 2002 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<%
Dim objCmd, objRs

Set objCmd = Server.CreateObject("ADODB.Command")

'[0]Command���� ����(Connection������Ʈ���� Ŀ�ǵ带 ������ �� �ĺ�)
objCmd.Name = "objCmdName"

'[1]Command ������Ʈ�� ���Ӹ��� ����(OLE DB���Ṯ�ڿ� �ۼ�).
objCmd.ActiveConnection = "Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)"

'[2]CommnadText �Ӽ��� ���� Command�� ������ ����(�⺻�� : 8(adCmdUnknown))
objCmd.CommandType = 8

'[3]Command ���ڿ��� ����.
objCmd.CommandText = "Select * From Memo"

'[4]Command ���� ��� �ð��� ����.(�⺻�� : 30��)
objCmd.CommandTimeout = 10

'[5]Command ���ڿ� �Ǵ� ��Ʈ��(Stream)�� ��Ÿ���� GUID�� ����.
Response.Write(objCmd.Dialect&"<br>")

'[6]Command�� �����ϱ� ���� SQL���� �������� ������ ����(True | False)
objCmd.Prepared = True

'[7]Parameter ������Ʈ�� ����(Append�޼ҵ�� �Բ� ���)
'Set objPara = Command������Ʈ.CreateParameter([Name],[Type],[Direction],[Size],[Value])
Set objPara = objCmd.CreateParameter("PrmID",3,1,1)
'�Ķ���͸� �߰�
objCmd.Parameters.Append objPara

'[8]Command�� ����
'[Set objRs =] Command������Ʈ.Execute([Records],[Para],[Type])
Set objRs = objCmd.Execute

'[9]Command ������Ʈ�� ����(0:����,1:����,...)
Response.Write(objCmd.State)

Response.Write("<div align=center><table border=1>")
Do Until objRs.EOF
	Response.Write("<tr>")
	For i = 0 To objRs.Fields.Count - 1
		Response.Write("<td>" & objRs(i) & "</td>")
	Next
	Response.Write("</tr>")
	objRs.MoveNext
Loop
Response.Write("</table></div>")

'[10] Command�� ���
objCmd.Cancel

objRs.Close

Set objRs = Nothing
Set objCmd = Nothing
%>