<%
'--------------------------------------------------
' Title : Basic ����
' Program Name : boardmodify.asp
' Program Description : �����ϱ� �� == �۾��� ���� ���
' QueryString : �ݵ�� view.asp���� Num=1 ������ ����
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
Option Explicit
Dim Num: Num = Request("Num")
Dim objCon: Dim objRs: Dim strSql
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open(Application("ConnectionString"))
Set objRs = Server.CreateObject("ADODB.RecordSet")
strSql = "Select * From Basic Where Num = " & Num
objRs.Open strSql, objCon
If objRs.BOF Or objRs.EOF Then
	Response.Write("������ �������� �߸��Ǿ����ϴ�.")
Else
	Call ShowModifyForm(objRs)
End If
objRs.Close(): objCon.Close()
Set objRs = Nothing: Set objCon = Nothing
%>
<%Sub ShowModifyForm(objRs)%>
<form action="./boardmodifyprocess.asp" method="post">
<pre>
<input type="hidden" name="Num" value="<%=Num%>">
�̸� : <input type="text" name="Name" value="<%=objRs("Name")%>">
�̸��� : <input type="text" name="Email" value="<%=objRs("Email")%>">
Ȩ������ : <input type="text" name="Homepage" value="<%=objRs("Homepage")%>">
���� : <input type="text" name="Title" value="<%=objRs("Title")%>">
���� : <textarea name="Content" cols="40" rows="10"><%=objRs("Content")%></textarea>
���ڵ� : <select name="Encoding"><option value="Text">Text</option><option value="HTML">HTML</option></select>
�н����� : <input type="password" name="Password">
<input type="submit" value="�ۼ� �Ϸ�">
</pre>
</form>
<%End Sub%>