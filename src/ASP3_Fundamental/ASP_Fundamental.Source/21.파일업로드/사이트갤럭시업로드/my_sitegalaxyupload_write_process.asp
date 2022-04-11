<%
' -----------------------------------------------------------
' Title : Using SiteGalaxyUpload Component
' Program Name : my_sitegalaxyupload_write_process.asp
' Description : ���Ͽø��� ó�� ���μ���
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: 
' Support: 
' -----------------------------------------------------------
%>
<%
'������ ����� ����
Option Explicit
Dim  objUploadForm, objFSO, FileName, strDirectory, AttachFile, FileSize
Dim strName, strExt, strFileName, tempFileName, i

'��ũ���� Ÿ�� �ƿ�
Server.ScriptTimeout = 900    

'����Ʈ������ ���ε� ������Ʈ�� �ν��Ͻ� ����
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form") 

'FSO ��ü�� �ν��Ͻ� ����
Set objFSO = CreateObject("Scripting.FileSystemObject")

'���ϸ� ����
FileName = objUploadForm("FileName") 

'���ϸ��� ���� ����
FileName = Trim(FileName)        

IF Len(FileName) > 0 Then '÷���� ������ ������
'������ ������ ����Ǵ� ��ġ���� : ���� ����� files����, files������ �̸� �����־� ������ �ȳ�.
strDirectory = Server.MapPath(".") + "\files\" 
Response.Write "[1] ������ ���� ���� ��ġ : " & strDirectory & "<br>"

AttachFile = objUploadForm("FileName").FilePath    '����� ������ ��� ����(Ŭ���̾�Ʈ ���� ��ġ)   
Response.Write "[2] ���ε� ������ ��ġ : " & AttachFile & "<br>"

FileSize = objUploadForm("FileName").Size '���� ������         
Response.Write "[3] ���ε� ������ ũ�� : " & FileSize & "<br>"

FileName = Mid(AttachFile, InstrRev(AttachFile, "\")+1) '���ϸ�.Ȯ���� �� ����!
Response.Write "[4] ���ϸ�.Ȯ���� : " & FileName & "<br>"
  
If InstrRev(FileName,".") Then
	strName = Mid(FileName, 1, InstrRev(FileName, ".")-1) '���ϸ� ����
Else
	strName = FileName
End If
Response.Write "[5] ���ϸ� : " & strName & "<br>"

strExt = Mid(FileName, Instr(FileName,".")+1) 'Ȯ���� ����
Response.Write "[6] Ȯ���� : " & strExt & "<br>"

'�ߺ��� ������ ������ �̸� �ڿ� ���ڸ� ����.
i = 0
Do While objFSO.FileExists(strDirectory & FileName)

	i = i + 1
	tempFileName = strName & "(" & i & ")"

	If strExt <> "" Then
		FileName = tempFileName & "." & strExt
	Else
		FileName = tempFileName
	End If

Loop

strFileName = strDirectory & FileName '������ ������ ��ο� ���� ����

Response.Write "[7] ������ ������ ��ο� ���ϸ�.Ȯ���� : " & strFileName & "<br>"

objUploadForm("FileName").SaveAs strFileName '������ ������ ����

End If 

Set objFSO = Nothing
Set objUploadForm = Nothing

Response.Write "[8] " & strFileName & "�� ����Ǿ����ϴ�."
%>
