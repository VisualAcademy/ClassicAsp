<%
Dim FileName '//�Ѱ��� �� ���ϸ�
Dim strDir '//������ ���� ���
Dim strName '//������ ���ϸ� : test.gif -> test
Dim strExt '//������ Ȯ���� : test.gif -> gif
'[0]����Ʈ������ ���ε� ������Ʈ�� �ν��Ͻ� ����
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form")
'[1]�Ѿ�� ������ ��ü ��� �� ��� : FilePath �Ӽ�
Response.Write("Ŭ���̾�Ʈ ��ü ��� : " & objUploadForm("FileName").FilePath & "<br>")
'[2]�Ѿ�� ������ ũ��(����Ʈ ����) : Size �Ӽ�
Response.Write("ũ�� : " & objUploadForm("FileName").Size) 
FileName = objUploadForm("FileName")'//Ŭ���̾�Ʈ ��ΰ� �ޱ�:FilePath
FileName = Mid(FileName, InStrRev(FileName, "\") + 1)'//���ϸ�и�
strName = Left(FileName, InStrRev(FileName, ".") - 1)'���ϸ��ϱ�
strExt = Mid(FileName, InStrRev(FileName, ".") + 1)'Ȯ���ڱ��ϱ�
'[!]���ϸ� �ߺ� ó�� : ���ϸ� �ڿ� ������� ��ȣ�� �ű�
Dim objFso
Set objFso = Server.CreateObject("Scripting.FileSystemObject")
'���ϸ� �ڿ� ��ȣ ���̱�
i = 0
Do While objFSO.FileExists(Server.MapPath(".") & "\files\" + FileName)
	i = i + 1
	tempFileName = strName & "(" & i & ")"
	If strExt <> "" Then
		FileName = tempFileName & "." & strExt
	Else
		FileName = tempFileName
	End If
Loop
'[!]������ ���� ��ο� ���ϸ�
strDir = Server.MapPath(".") & "\files\" & FileName	'//
'[3]���ε��ϱ� : SaveAs() �޼��� ���
objUploadForm("FileName").SaveAs strDir				'//
%>












