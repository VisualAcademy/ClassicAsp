<%
	'[1] ��������
	Option Explicit
	Dim objFso, objFile

	'[2] FSO��ü�� �ν��Ͻ� ����
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")

	'[3] ���� ��ü ���� : �ڵ鸵 �о���� : �б������� 1��
	Set objFile = objFso.OpenTextFile("C:\Temp\test1.txt",1)

	'[4] ���� �о���� : ReadLine�޼ҵ�� ���پ� �о�´�.
	Response.Write objFile.ReadLine & "<br>"
	Response.Write objFile.ReadLine & "<br>"
	Response.Write objFile.ReadLine & "<br>"

	'[5] ��ü ����
	Set objFile = Nothing
	Set objFso = Nothing
%>
