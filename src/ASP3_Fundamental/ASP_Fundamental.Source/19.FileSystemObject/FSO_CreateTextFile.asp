<%
	'[1]���� ����
	Option Explicit
	Dim objFSO   'FSO��ü�� �ν��Ͻ� ����
	Dim objFile   'FSO��ü�� File��ü�� �ڵ鸵 ����

	'[2]FSO ��ü�� �ν��Ͻ�(��ü)�� �����Ѵ�.(���� ���� ��ü�� ���赵�� Scripting.FileSystemObject)
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'[3]���ϻ��� : CreateTextFile("�����̸��װ��",���������,�����ڵ���������)
	Set objFile = objFSO.CreateTextFile("C:\test1.txt",true,false)
	Set objFile = objFSO.CreateTextFile("C:\test2.txt",true,false)
	Set objFile = objFSO.CreateTextFile("C:\test3.txt",true,false)

	'���� �������� üũ : FileExists -> �����ϸ� True, �׷��������� False�� ��ȯ �޼ҵ�.
	If objFSO.FileExists("C:\test1.txt") Then
		Response.Write("test1.txt ������ ����������ϴ�.<br>")
	End If
	If objFSO.FileExists("C:\test2.txt") Then
		Response.Write("test2.txt ������ ����������ϴ�.<br>")
	End If
	If objFSO.FileExists("C:\test3.txt") Then
		Response.Write("test3.txt ������ ����������ϴ�.<br>")
	End If

	'[4]FSO��ü�� �ν��Ͻ� ����(ASP�� �˾Ƽ� ���� ����������, �⺻������ �ۼ��ض�...)
	Set objFile = Nothing
	Set objFSO = Nothing
%>
