<%
'[1] ���� ����
Option Explicit
Dim objCon
Dim strSql
Dim objRs
'[2] �����ͺ��̽��� �ν��Ͻ� ���� : Connection ������Ʈ
Set objCon = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.Recordset")'//
'[3] �����ͺ��̽� ����
objCon.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")
'[4] SQL�� ����
strSql = "Select * From Memo"
objRs.Open strSql, objCon'//
If objRs.BOF Or objRs.EOF Then
	Response.Write("�Էµ� �����Ͱ� �����ϴ�.")
Else
	Call ShowRecordSet(objRs)
End If
'[5] �����ͺ��̽� �ݱ�
objRs.Close()'//
objCon.Close()
'[6] �����ͺ��̽��� �ν��Ͻ� ����
Set objRs = Nothing'//
Set objCon = Nothing 
%>
<%
Sub ShowRecordSet(objRs)
	While Not objRs.EOF
		Response.Write(objRs.Fields("Name") & Space(1) & _
		objRs("Title") & Space(1) & _
		objRs.Fields(4) & "<br>")
		objRs.MoveNext()
	WEnd
End Sub 
%>
