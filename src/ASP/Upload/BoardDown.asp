<%
'//���� ���� �ٿ�ε� ������ : BoardDown.asp
Call Main()
Sub Main()
     strFileName = Request("strFileName")
     If strFileName = "" Then
          Call RedirFail()
          Exit Sub     
     Else
		Call UpdateDownCount(strFileName)'//�ٿ�ε� Ƚ�� ���� �Լ�ȣ��
          Call RedirectFile(strFileName)'//�ٿ�ε� ����
     End If
End Sub

Sub UpdateDownCount(strFileName)'//�ٿ�ε� Ƚ�� ���� ����:��ȸ�������ͺ��
	Dim objCon: Dim objRs
	Set objCon = Server.CreateObject("ADODB.Connection")
	objCon.Open(Application("ConnectionString"))
	strSql = "Select FileName From Upload Where FileName = '" & strFileName & "'"
	Set objRs = objCon.Execute(strSql)
	If objRs.BOF Or objRs.EOF Then
	Else
		strSql = "Update Upload Set DownCount = DownCount + 1 Where FileName = '" & strFileName & "'"
		objCon.Execute(strSql)
	End If	
	objCon.Close()
	Set objCon = Nothing
End Sub

Sub RedirectFile(strFileName)
     '������ ����� ������ ��ġ�� �̸�
     strFilePath = Server.MapPath(".") + "\files\" & strFileName

     Set objFSO = Server.CreateObject("SiteGalaxyUpload.FileSystemObject")
     Set objFile = objFSO.OpenBinaryFile(strFilePath, 1, False)
     
     Response.Clear
     Response.ContentType = "application/octet-stream"
     Response.AddHeader "Content-disposition", "attachment; filename=" & strFileName
     Response.BinaryWrite objFile.ReadAll
     Response.End
     
     Set objFile = Nothing
     Set objFSO = Nothing

End Sub


Sub RedirFail()

     Response.Redirect("./boardlist.asp")

End Sub

%> 