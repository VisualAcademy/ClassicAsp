<%
'1. ����Ʈ������ ���ε� ������Ʈ ��ġ
'2. �۾��� �� ����������...
'	<form> �±׿�
'	enctype="multipart/form-data" �Ӽ� �߰�
'3. �۾��� ó�� ����������... Request��ü�� ��� ���ϹǷ�
'	����Ʈ������ ���ε� ������Ʈ�� �ν��Ͻ� ����
'4. ������ �ʵ带 �ν��Ͻ� �̸����� �ޱ�
'5. �Ѱ��� �� ��ü ��� �� �ޱ�
'6. �Ѱ��� �� ������ ũ�� �� �ޱ�
'7. ��ü ��ο��� ���ϸ� �и�
'8. ���ϸ��� �������ϸ�� Ȯ���� �и�
'9. ���ε��� ���� ����

'�ڷ�� 3 : ����Ʈ ������ ���ε� ������Ʈ�� �ν��Ͻ� ����-----------------
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form")
'---------------------------------------------------------------

'�ڷ�� 4 : Request��ü ��ſ� objUploadForm��ü�� ����---------------
'���� ����
Dim objCon: Dim strSql
Dim Name: Name = objUploadForm("Name")
Dim Email: Email = objUploadForm("Email")
Dim Title: Title = objUploadForm("Title")
Dim PostIP: PostIP = Request.ServerVariables("REMOTE_HOST")
Dim Content: Content = objUploadForm("Content")
Dim Password: Password = objUploadForm("Password")
Dim Encoding: Encoding = objUploadForm("Encoding")
Dim Homepage: Homepage = objUploadForm("Homepage")
'Dim FileName: FileName = objUploadForm("FileName")
'Dim FileSize: FileSize = 100
'---------------------------------------------------------------

'�ڷ�� 5 : �տ��� �Ѱ��� �� ������ ��ü ��ΰ� �ޱ�---------------------
Dim FullFilePath
FullFilePath = objUploadForm("FileName").FilePath
'---------------------------------------------------------------

'�ڷ�� 6 : �տ��� �Ѱ��� �� ������ ũ��(����Ʈ ����)--------------------
Dim FileSize
FileSize = objUploadForm("FileName").Size
'---------------------------------------------------------------

'�ڷ�� 7 : ��ü ��ο��� ���ϸ�� Ȯ���� �и�--------------------
Dim FileName
FileName = Mid(FullFilePath, InStrRev(FullFilePath, "\") + 1)
'---------------------------------------------------------------

'�ڷ�� 8 : ���ϸ��� ���� ���ϸ�� Ȯ���� ���� �и�-------------------
Dim strName	'���� ���ϸ�
strName = Left(FileName, InStrRev(FileName, ".") - 1)
Dim strExt	'���� Ȯ����
strExt = Mid(FileName, InStrRev(FileName, ".") + 1)
'---------------------------------------------------------------

'�ڷ�� 9 : ���� ���ε� �� ���� ������ ����-------------------
Dim strDir
strDir = Server.MapPath(".") & "\files\"  '�ݵ�� ���� ���� �ʿ�
'---------------------------------------------------------------

'�ڷ�� 10 : ���ϸ� �ߺ�ó��-------------------
Dim objFso
Set objFso = Server.CreateObject("Scripting.FileSystemObject")
i = 0
Do While objFSO.FileExists(strDir + FileName)
	i = i + 1
	tempFileName = strName & "(" & i & ")"
	If strExt <> "" Then
		FileName = tempFileName & "." & strExt
	Else
		FileName = tempFileName
	End If
Loop
'---------------------------------------------------------------

'�ڷ�� 11 : ���� ����-------------------
objUploadForm("FileName").SaveAs strDir & FileName
'---------------------------------------------------------------

'[!] ��������ǥ ó��
Name = Replace(Name, "'", "''")
Title = Replace(Title, "'", "''")
Content = Replace(Content, "'", "''")
'Ŀ�ؼ� ��ü ����
Set objCon = Server.CreateObject("ADODB.Connection")
'Ŀ�ؼ� ��ü ����
objCon.Open(Application("ConnectionString"))
'�̸� ������ SQL�� ���� : �����ѵ����� -> " & ���� & "
strSql = "Insert Upload(Name,Email,Title,PostDate,PostIP,Content,Password,ReadCount,Encoding,Homepage,ModifyDate,ModifyIP,FileName,FileSize,DownCount) Values('" & Name & "','" & Email & "','" & Title & "',GetDate(),'" & PostIP & "','" & Content & "',	'" & Password & "',0,'" & Encoding & "','" & Homepage & "',Null,Null,'" & FileName & "'," & FileSize & ",0)"
'SQL�� ����
objCon.Execute(strSql)
'Ŀ�ؼ� ��ü �ݱ�
objCon.Close()
'Ŀ�ؼ� ��ü ����
Set objCon = Nothing
'����Ʈ �������� �̵�
Response.Redirect("BoardList.asp")
%>
<center>�����Ͱ� �ԷµǾ����ϴ�.</center>