<%
'--------------------------------------------------
' Title : MyUpload Education Version
' Program Name : upload_write_process.asp
' Program Description : �Է������� ���۵� �� Insert������ ����
' Include Files : upload_function.asp
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!-- HTML�±� �� SQL Ư������  ���� ���� ���� ����� �����Լ� ȣ�� -->
<!-- #include file="./upload_function.asp" -->

<%
'Option Explicit   '�� ������ �ּ�ó���� ������ ��Ŭ��� ���Ͽ��� ���������� ���߱⶧��(��������)

'���� ����
Dim Name, Email, Title, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Content, objDB, strSQL
Dim Num, Ref, Step, RefOrder, Number, objRS

' -------------------- �ڷ�ǿ����� �߰� --------------------
Dim  objUploadForm, objFSO, FileName, strDirectory, AttachFile, FileSize, strSQL2
Dim strName, strExt, strFileName, tempFileName, i

'��ũ���� Ÿ�� �ƿ�
Server.ScriptTimeout = 900    

'����Ʈ������ ���ε� ������Ʈ�� �ν��Ͻ� ����
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form") 
' -------------------- �ڷ�ǿ����� �߰� --------------------

'����Ʈ������ ���ε� ������Ʈ�� ��ü�� ����ؼ� �Ѿ�� ������(name, value)�� �޴� ����.
Num = objUploadForm("Num")
Name = objUploadForm("Name")
Email = objUploadForm("Email")
Title = objUploadForm("Title")
PostDate = Date()
ModifyDate = Date() 'ó�� �۾����� �ش� ��¥ �Է�.
ReadCount = 0
PostIP = Request.ServerVariables("REMOTE_HOST")
ModifyIP = Request.ServerVariables("REMOTE_HOST") 'ó�� �۾����� �ش� IP �Է�.
Encoding = objUploadForm("Encoding")
Passwd = objUploadForm("Password")
Content = objUploadForm("Content")

'HTML�±� �� SQL Ư������  ���� ����.-----------------------------------
	Title = ConvertToHTML(Title)   '���� HTML������
	Email = ConvertToHTML(Email)   'Email�� HTML������
	
	Title = ConvertToSQL(Title)   'Title�� ��������ǥ ó��
	Email = ConvertToSQL(Email)   'Email�� ��������ǥ ó��
	Content = ConvertToSQL(Content)   'Content�� ��������ǥ ó��

	'���� �� ���ڵ��� HTML���� Text������ ���� ���� ���� ����.
	Select Case Encoding

		Case "HTML"   'HTML�� ó���ؼ� �����ٶ�
	
			Content = CStr(Content)
		
		Case "Text"   'HTML �ҽ��� �״�� ������ ��

			Content = CStr(Content)
			Content = ConvertToHTML(Content)   

	End Select
'HTML�±� �� SQL Ư������  ���� ����.-----------------------------------

'�����Ͱ� ����� �Ѿ�Դ��� �����
Response.Write(Name & "," & Email & "," & Title & "," & PostDate & "," & ModifyDate & "," & ReadCount & "," & PostIP & "," & ModifyIP & "," & Encoding & "," & Passwd& "," & Content) & "<br>"


' -------------------- Q & A �Խ��ǿ����� �߰� --------------------
'�����ͺ��̽�(connection��ü)�� �ν��Ͻ� ����
Set objDB = Server.Createobject("ADODB.Connection")

'�����ͺ��̽� ����
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")
 
 '�Խ����� �� ���ڵ� ���� Number��� �̸����� ��ȯ
strSQL = "Select Count(Num) As Number From my_upload"

'���ڵ�� ��ü�� �ν��Ͻ� ����
Set objRS = Server.Createobject("ADODB.RecordSet")

'���ڵ���� �μ��� ������ ��...(���߿� RecordSet��ü �ð��� �ڼ��� �ٷ�.)
objRS.Open strSQL, objDB, 2, 1, 1

'If objRS.BOF OR objRS.EOF Then
'	Response.Write "�����Ͱ� ������."
'End IF

If objRS.BOF OR objRS.EOF Then '��ȣ ����
	Number = 1
Else 
	Number = objRS("Number") + 1
End If

'Response.Write Number
'Response.Write "Ref :" & Ref & "<br>"

If Num <> "" Then '�� �亯������

	Ref = objUploadForm("Ref")
	Step = objUploadForm("Step")
	RefOrder = objUploadForm("RefOrder")

	strSQL = "Update my_upload Set RefOrder = RefOrder + 1 Where Ref = " & Ref & " AND RefOrder > " & RefOrder
	Response.Write strSQl & "<br>"


	objDB.Execute(strSQL)

	Step = Step + 1
	RefOrder = RefOrder + 1

Else ' ÷ �۾�����...

Ref = Number
Step = 0
RefOrder = 0

End If
' -------------------- Q & A �Խ��ǿ����� �߰� --------------------
'Response.End
' -------------------- �ڷ�ǿ����� �߰� --------------------

'FSO ��ü�� �ν��Ͻ� ����
Set objFSO = CreateObject("Scripting.FileSystemObject")

'���ϸ� ����
FileName = objUploadForm("FileName") 

If FileName = "" Then
	FileName = Null
	FileSize = 0
Else

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

strExt = Mid(FileName, InstrRev(FileName,".")+1) 'Ȯ���� ����
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



Response.Write "[8] " & strFileName & "�� ����Ǿ����ϴ�."
' -------------------- �ڷ�ǿ����� �߰� --------------------

End If

'�����ͺ��̽� ó��(�Է�) ���� �ۼ� : ���ϸ�� ����ũ�⸦ �����ͺ��̽��� ����
strSQL2 = "Insert my_upload(Name, Email, Title, FileName, FileSize, PostDate, ModifyDate, ReadCount, PostIP, ModifyIP, Encoding, Passwd, Ref, Step, RefOrder, Content) Values "
strSQL2 = strSQL2 & "('" & Name & "','" & Email & "','" & Title & "','" & FileName & "'," 
strSQL2 = strSQL2 & FileSize & ",'" & PostDate & "','" 
strSQL2 = strSQL2 & ModifyDate & "'," & ReadCount & ",'" & PostIP & "','" & ModifyIP & "','"
strSQL2 = strSQL2 & Encoding & "','" & Passwd & "'," & Ref & "," & Step & "," & RefOrder & ",'" & Content & "')"
        
'���� ���� �����
Response.Write strSQL2 & "<br>"

'�����ͺ��̽� ó��(�Է�) ����
objDB.Execute(strSQL2)

%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<center>
		<br>������ ���̽��� �ԷµǾ����ϴ�. <br>
		<a href="./upload_list.asp">���� ����Ʈ�� �̵�</a>
		</center>
	</BODY>
</HTML>

<%
'�����ͺ��̽� �ݱ�
objDB.Close
'�����ͺ��̽� ��ü�� �ν��Ͻ� ���� : ������Ʈ ��ü ������ �ݵ�� ��� ó���� ������ �Ѵ�.
Set objFSO = Nothing
Set objUploadForm = Nothing
Set objDB = Nothing
%>
