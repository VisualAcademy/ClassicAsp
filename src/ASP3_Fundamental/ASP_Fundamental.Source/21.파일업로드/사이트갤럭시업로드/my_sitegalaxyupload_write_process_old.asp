<%
' -----------------------------------------------------------
' Title : Using SiteGalaxyUpload Component
' Program Name : my_sitegalaxyupload_write_process.asp
' Description : ���Ͽø��� ó�� ���μ���(�ߺ� ���� 1�� ó��)
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
' -----------------------------------------------------------
%>
<%
'������ ����� ����
Option Explicit
Dim  objUploadForm, objFSO, FileName, strDirectory, AttachFile, FileSize, strName, strExt, strFileName, countFileName, bExist

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
strDirectory = Server.MapPath(".") + "\file\" '������ ������ ����Ǵ� ��ġ����
Response.Write "[1] ������ ���� ���� ��ġ : " & strDirectory & "<br>"

AttachFile = objUploadForm("FileName").FilePath    '����� ������ ��� ����(Ŭ���̾�Ʈ ���� ��ġ)   
Response.Write "[2] ���ε� ������ ��ġ : " & AttachFile & "<br>"

FileSize = objUploadForm("FileName").Size '���� ������         
Response.Write "[3] ���ε� ������ ũ�� : " & FileSize & "<br>"

FileName = Mid(AttachFile, InstrRev(AttachFile, "\")+1) '���ϸ�.Ȯ���� �� ����!
Response.Write "[4] ���ϸ�.Ȯ���� : " & FileName & "<br>"
  
strName = Mid(FileName, 1, Instr(FileName, ".")-1) '���ϸ� ����
Response.Write "[5] ���ϸ� : " & strName & "<br>"

strExt = Mid(FileName, Instr(FileName,".")+1) 'Ȯ���� ����
Response.Write "[6] Ȯ���� : " & strExt & "<br>"

bExist = True '���� ������ �����Ѵٰ� ����
strFileName = strDirectory&strName&"."&strExt '������ ������ ��ο� ���� ����
Response.Write "[7] ������ ������ ��ο� ���ϸ�.Ȯ���� : " & strFileName & "<br>"



countFileName = 0 '���� �����϶� ���ϸ�ڿ� ���� ����
        
IF bExist = True Then
	IF (objFSO.FileExists(strFileName)) Then '���� ������ �����ϸ�
		countFileName = countFileName+1
		FileName = strName&"_"&countFileName&"."&strExt
		strFileName = strDirectory&FileName '���� ������ ����� ��ο� ���ϸ�
	Else
		bExist = False
	End IF
End IF
Response.Write "[8] ���������� ����� ���ϸ� : " & strFileName & "<br>"

objUploadForm("FileName").SaveAs strFileName '������ ������ ����

End If 

Set objFSO = Nothing
Set objUploadForm = Nothing

Response.Write "<br>" & strFileName & "�� ����Ǿ����ϴ�."
%>
