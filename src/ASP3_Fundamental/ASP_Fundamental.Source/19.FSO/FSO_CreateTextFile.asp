<%
	'[1]���� ����
	Option Explicit
	Dim objFso   'FSO��ü�� �ν��Ͻ� ����
	Dim objFile   'FSO��ü�� File��ü�� �ڵ鸵 ����

	'[2]FSO ��ü�� �ν��Ͻ�(��ü)�� �����Ѵ�.(���� ���� ��ü�� ���赵�� Scripting.FileSystemObject)
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")

	'[3]���ϻ��� : CreateTextFile("�����̸��װ��",���������,�����ڵ���������)
	Set objFile = objFso.CreateTextFile("C:\Temp\test1.txt",true,false)
	Set objFile = objFso.CreateTextFile("C:\Temp\test2.txt",true,false)
	Set objFile = objFso.CreateTextFile("C:\Temp\test3.txt",true,false)

	'���� �������� üũ : FileExists -> �����ϸ� True, �׷��������� False�� ��ȯ �޼ҵ�.
	If objFso.FileExists("C:\Temp\test1.txt") Then
		Response.Write("test1.txt ������ ����������ϴ�.<br>")
	End If
	If objFso.FileExists("C:\Temp\test2.txt") Then
		Response.Write("test2.txt ������ ����������ϴ�.<br>")
	End If
	If objFso.FileExists("C:\Temp\test3.txt") Then
		Response.Write("test3.txt ������ ����������ϴ�.<br>")
	End If

	'[4]FSO��ü�� �ν��Ͻ� ����(ASP�� �˾Ƽ� ���� ����������, �⺻������ �ۼ��ض�...)
	Set objFile = Nothing
	Set objFso = Nothing
%>
